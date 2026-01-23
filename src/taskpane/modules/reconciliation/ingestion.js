/**
 * OOXML Reconciliation Pipeline - Ingestion
 * 
 * Parses OOXML paragraph content into a run model with character offsets.
 * Handles track changes, hyperlinks, and basic container elements.
 */

import { NS_W, RunKind } from './types.js';

/**
 * Ingests OOXML paragraph content and builds a run-aware text model.
 * 
 * @param {string} ooxmlString - The OOXML string to parse
 * @returns {import('./types.js').IngestionResult} The ingestion result
 */
export function ingestOoxml(ooxmlString) {
    const runModel = [];
    let acceptedText = '';

    if (!ooxmlString) {
        return { runModel, acceptedText, pPr: null };
    }

    try {
        const parser = new DOMParser();
        const doc = parser.parseFromString(ooxmlString, 'application/xml');

        // Check for parse errors
        const parseError = doc.getElementsByTagName('parsererror')[0];
        if (parseError) {
            console.error('OOXML parse error:', parseError.textContent);
            return { runModel, acceptedText, pPr: null };
        }

        // Find the paragraph element
        const paragraph = doc.getElementsByTagNameNS(NS_W, 'p')[0];
        if (!paragraph) {
            // Try without namespace (some OOXML comes with default namespace)
            const pElements = doc.getElementsByTagName('w:p');
            if (pElements.length === 0) {
                console.warn('No paragraph found in OOXML');
                return { runModel, acceptedText, pPr: null };
            }
        }

        const pElement = paragraph || doc.getElementsByTagName('w:p')[0];

        // Extract paragraph properties
        const pPr = pElement.getElementsByTagNameNS(NS_W, 'pPr')[0] ||
            pElement.getElementsByTagName('w:pPr')[0] || null;

        // Process the paragraph content
        let currentOffset = 0;
        const result = processNodeRecursive(pElement, currentOffset, runModel);
        acceptedText = result.text;

        return { runModel, acceptedText, pPr };

    } catch (error) {
        console.error('Error ingesting OOXML:', error);
        return { runModel, acceptedText, pPr: null };
    }
}

let containerIdCounter = 0;

/**
 * Recursively processes nodes and builds the run model.
 * 
 * @param {Node} node - The node to process
 * @param {number} currentOffset - Current character offset
 * @param {Array} runModel - Run model array to populate
 * @returns {{ offset: number, text: string }}
 */
function processNodeRecursive(node, currentOffset, runModel) {
    let localOffset = currentOffset;
    let text = '';

    for (const child of node.childNodes) {
        const nodeName = child.nodeName;

        // Skip: paragraph properties, proofing markers
        if (nodeName === 'w:pPr') continue;
        if (nodeName === 'w:proofErr') continue;

        // Content Control (SDT)
        if (nodeName === 'w:sdt') {
            const containerId = `sdt_${containerIdCounter++}`;
            const sdtPr = child.getElementsByTagNameNS(NS_W, 'sdtPr')[0] ||
                child.getElementsByTagName('w:sdtPr')[0];
            const sdtContent = child.getElementsByTagNameNS(NS_W, 'sdtContent')[0] ||
                child.getElementsByTagName('w:sdtContent')[0];

            runModel.push({
                kind: RunKind.CONTAINER_START,
                containerKind: ContainerKind.SDT,
                containerId,
                propertiesXml: sdtPr ? new XMLSerializer().serializeToString(sdtPr) : '',
                startOffset: localOffset,
                endOffset: localOffset,
                text: ''
            });

            if (sdtContent) {
                const result = processNodeRecursive(sdtContent, localOffset, runModel);
                localOffset = result.offset;
                text += result.text;
            }

            runModel.push({
                kind: RunKind.CONTAINER_END,
                containerKind: ContainerKind.SDT,
                containerId,
                startOffset: localOffset,
                endOffset: localOffset,
                text: ''
            });
            continue;
        }

        // Smart Tag
        if (nodeName === 'w:smartTag') {
            const containerId = `smartTag_${containerIdCounter++}`;

            runModel.push({
                kind: RunKind.CONTAINER_START,
                containerKind: ContainerKind.SMART_TAG,
                containerId,
                propertiesXml: serializeAttributes(child),
                startOffset: localOffset,
                endOffset: localOffset,
                text: ''
            });

            const result = processNodeRecursive(child, localOffset, runModel);
            localOffset = result.offset;
            text += result.text;

            runModel.push({
                kind: RunKind.CONTAINER_END,
                containerKind: ContainerKind.SMART_TAG,
                containerId,
                startOffset: localOffset,
                endOffset: localOffset,
                text: ''
            });
            continue;
        }

        // Skip deleted content (w:del) - not part of accepted text
        if (nodeName === 'w:del') {
            const deletionEntry = processDeletion(child, localOffset);
            if (deletionEntry) {
                runModel.push(deletionEntry);
            }
            continue;
        }

        // Bookmarks: preserve but no text contribution
        if (nodeName === 'w:bookmarkStart' || nodeName === 'w:bookmarkEnd') {
            runModel.push({
                kind: RunKind.BOOKMARK,
                nodeXml: new XMLSerializer().serializeToString(child),
                startOffset: localOffset,
                endOffset: localOffset,
                text: ''
            });
            continue;
        }

        // Insertions: content IS part of accepted text
        if (nodeName === 'w:ins') {
            const result = processNodeRecursive(child, localOffset, runModel);
            localOffset = result.offset;
            text += result.text;
            continue;
        }

        // Hyperlinks: use container tokens for better preservation
        if (nodeName === 'w:hyperlink') {
            const containerId = `hyperlink_${containerIdCounter++}`;
            const rId = child.getAttribute('r:id') || '';
            const anchor = child.getAttribute('w:anchor') || '';

            runModel.push({
                kind: RunKind.CONTAINER_START,
                containerKind: 'hyperlink',
                containerId,
                propertiesXml: JSON.stringify({ rId, anchor }),
                startOffset: localOffset,
                endOffset: localOffset,
                text: ''
            });

            const result = processNodeRecursive(child, localOffset, runModel);
            localOffset = result.offset;
            text += result.text;

            runModel.push({
                kind: RunKind.CONTAINER_END,
                containerKind: 'hyperlink',
                containerId,
                startOffset: localOffset,
                endOffset: localOffset,
                text: ''
            });
            continue;
        }

        // Standard runs (w:r)
        if (nodeName === 'w:r') {
            const runEntry = processRun(child, localOffset);
            if (runEntry && runEntry.text) {
                runModel.push(runEntry);
                localOffset += runEntry.text.length;
                text += runEntry.text;
            }
            continue;
        }
    }

    return { offset: localOffset, text };
}

function serializeAttributes(element) {
    return Array.from(element.attributes)
        .map(attr => `${attr.name}="${attr.value}"`)
        .join(' ');
}

/**
 * Processes a standard w:r run element.
 * 
 * @param {Element} runElement - The w:r element
 * @param {number} startOffset - Starting character offset
 * @returns {import('./types.js').RunEntry|null}
 */
function processRun(runElement, startOffset) {
    // Extract run properties (formatting)
    const rPr = runElement.getElementsByTagNameNS(NS_W, 'rPr')[0] ||
        runElement.getElementsByTagName('w:rPr')[0];
    const rPrXml = rPr ? new XMLSerializer().serializeToString(rPr) : '';

    // Extract text content
    let text = '';
    const textNodes = runElement.getElementsByTagNameNS(NS_W, 't');
    if (textNodes.length === 0) {
        // Try without namespace
        const tNodes = runElement.getElementsByTagName('w:t');
        for (const t of tNodes) {
            text += t.textContent || '';
        }
    } else {
        for (const t of textNodes) {
            text += t.textContent || '';
        }
    }

    // Handle tabs and breaks
    const tabs = runElement.getElementsByTagNameNS(NS_W, 'tab');
    const breaks = runElement.getElementsByTagNameNS(NS_W, 'br');
    if (tabs.length > 0) text += '\t'.repeat(tabs.length);
    if (breaks.length > 0) text += '\n'.repeat(breaks.length);

    if (!text) return null;

    return {
        kind: RunKind.TEXT,
        text,
        rPrXml,
        startOffset,
        endOffset: startOffset + text.length
    };
}

/**
 * Processes a w:del deletion element.
 * Stores the deleted content but doesn't contribute to accepted text.
 * 
 * @param {Element} delElement - The w:del element
 * @param {number} offset - Current offset (unchanged)
 * @returns {import('./types.js').RunEntry|null}
 */
function processDeletion(delElement, offset) {
    const author = delElement.getAttribute('w:author') || '';

    // Extract text from w:delText elements
    let text = '';
    const delTexts = delElement.getElementsByTagNameNS(NS_W, 'delText');
    for (const dt of delTexts) {
        text += dt.textContent || '';
    }

    // Also look inside w:r elements for w:delText
    const runs = delElement.getElementsByTagNameNS(NS_W, 'r');
    for (const run of runs) {
        const innerDelTexts = run.getElementsByTagNameNS(NS_W, 'delText') ||
            run.getElementsByTagName('w:delText');
        for (const dt of innerDelTexts) {
            text += dt.textContent || '';
        }
    }

    if (!text) return null;

    return {
        kind: RunKind.DELETION,
        text,
        rPrXml: '',
        startOffset: offset,
        endOffset: offset, // Deletions don't advance offset
        author,
        nodeXml: new XMLSerializer().serializeToString(delElement)
    };
}

/**
 * Processes a w:hyperlink element.
 * 
 * @param {Element} hyperlinkElement - The w:hyperlink element
 * @param {number} startOffset - Starting character offset
 * @returns {import('./types.js').RunEntry}
 */
function processHyperlink(hyperlinkElement, startOffset) {
    const rId = hyperlinkElement.getAttribute('r:id') || '';
    const anchor = hyperlinkElement.getAttribute('w:anchor') || '';

    let text = '';
    const runs = hyperlinkElement.getElementsByTagNameNS(NS_W, 'r') ||
        hyperlinkElement.getElementsByTagName('w:r');

    for (const run of runs) {
        const textNodes = run.getElementsByTagNameNS(NS_W, 't') ||
            run.getElementsByTagName('w:t');
        for (const t of textNodes) {
            text += t.textContent || '';
        }
    }

    return {
        kind: RunKind.HYPERLINK,
        text,
        startOffset,
        endOffset: startOffset + text.length,
        rId,
        anchor,
        rPrXml: '',
        nodeXml: new XMLSerializer().serializeToString(hyperlinkElement)
    };
}

/**
 * Extracts numbering information from a paragraph element.
 * 
 * @param {Element} pElement - The w:p element
 * @returns {Object|null} Numbering context { numId, ilvl }
 */
export function detectNumberingContext(pElement) {
    const pPr = pElement.getElementsByTagNameNS(NS_W, 'pPr')[0] ||
        pElement.getElementsByTagName('w:pPr')[0];
    if (!pPr) return null;

    const numPr = pPr.getElementsByTagNameNS(NS_W, 'numPr')[0] ||
        pPr.getElementsByTagName('w:numPr')[0];
    if (!numPr) return null;

    const numIdEl = numPr.getElementsByTagNameNS(NS_W, 'numId')[0] ||
        numPr.getElementsByTagName('w:numId')[0];
    const ilvlEl = numPr.getElementsByTagNameNS(NS_W, 'ilvl')[0] ||
        numPr.getElementsByTagName('w:ilvl')[0];

    if (!numIdEl) return null;

    const numId = numIdEl.getAttribute('w:val');
    // Basic heuristic: numId 1 is usually bullet, 2 is usually numbered in Word defaults
    // This is a stub until we parse numbering.xml properly
    const type = numId === '1' ? 'bullet' : (numId === '2' ? 'numbered' : 'unknown');

    return {
        numId,
        ilvl: parseInt(ilvlEl?.getAttribute('w:val') || '0', 10),
        type
    };
}

/**
 * Ingests a table into a virtual grid model to handle merged cells.
 * 
 * @param {Element} tableNode - The w:tbl element
 * @returns {Object} The virtual grid model
 */
export function ingestTableToVirtualGrid(tableNode) {
    const tblGrid = tableNode.getElementsByTagNameNS(NS_W, 'tblGrid')[0] ||
        tableNode.getElementsByTagName('w:tblGrid')[0];
    const gridCols = tblGrid ? (tblGrid.getElementsByTagNameNS(NS_W, 'gridCol').length > 0 ?
        tblGrid.getElementsByTagNameNS(NS_W, 'gridCol') :
        tblGrid.getElementsByTagName('w:gridCol')) : [];
    const colCount = gridCols.length;

    const trElements = tableNode.getElementsByTagNameNS(NS_W, 'tr').length > 0 ?
        tableNode.getElementsByTagNameNS(NS_W, 'tr') :
        tableNode.getElementsByTagName('w:tr');
    const rowCount = trElements.length;

    // Initialize empty grid
    const grid = Array.from({ length: rowCount }, () =>
        Array.from({ length: colCount }, () => null)
    );
    const cellMap = new Map();

    // Track vertical merge continuations
    const vMergeOrigins = new Map(); // Key: colIdx -> { originRow, cell }

    for (let rowIdx = 0; rowIdx < trElements.length; rowIdx++) {
        const tr = trElements[rowIdx];
        const tcElements = tr.getElementsByTagNameNS(NS_W, 'tc').length > 0 ?
            tr.getElementsByTagNameNS(NS_W, 'tc') :
            tr.getElementsByTagName('w:tc');

        let gridCol = 0; // Current position in virtual grid

        for (let tcIdx = 0; tcIdx < tcElements.length; tcIdx++) {
            const tc = tcElements[tcIdx];
            const tcPr = tc.getElementsByTagNameNS(NS_W, 'tcPr')[0] ||
                tc.getElementsByTagName('w:tcPr')[0];

            // Skip to next available grid column (may be occupied by vMerge from above)
            while (gridCol < colCount && grid[rowIdx][gridCol] !== null) {
                gridCol++;
            }

            if (gridCol >= colCount) break; // Row is full

            // Parse gridSpan (horizontal merge)
            const gridSpanEl = tcPr ? (tcPr.getElementsByTagNameNS(NS_W, 'gridSpan')[0] ||
                tcPr.getElementsByTagName('w:gridSpan')[0]) : null;
            const colSpan = parseInt(gridSpanEl?.getAttribute('w:val') || '1', 10);

            // Parse vMerge (vertical merge)
            const vMergeEl = tcPr ? (tcPr.getElementsByTagNameNS(NS_W, 'vMerge')[0] ||
                tcPr.getElementsByTagName('w:vMerge')[0]) : null;
            const vMergeVal = vMergeEl?.getAttribute('w:val'); // "restart" or undefined (continue)
            const hasVMerge = vMergeEl !== null;

            let cell;

            if (hasVMerge && vMergeVal !== 'restart') {
                // This is a vMerge continuation - link to origin
                const origin = vMergeOrigins.get(gridCol);
                if (origin) {
                    origin.cell.rowSpan++;
                    cell = {
                        gridRow: rowIdx,
                        gridCol,
                        rowSpan: 0, // Continuation cells have rowSpan 0 in our logic
                        colSpan,
                        tcNode: tc,
                        blocks: [], // Content comes from origin
                        tcPrXml: serializeTcPr(tcPr),
                        isMergeOrigin: false,
                        isMergeContinuation: true,
                        mergeOrigin: origin.cell
                    };
                } else {
                    // Fallback for malformed XML
                    cell = createRegularCell(rowIdx, gridCol, colSpan, tc, tcPr);
                }
            } else {
                // Regular cell or vMerge="restart" (origin of vertical merge)
                const blocks = parseCellBlocks(tc);

                cell = {
                    gridRow: rowIdx,
                    gridCol,
                    rowSpan: 1,
                    colSpan,
                    tcNode: tc,
                    blocks,
                    tcPrXml: serializeTcPr(tcPr),
                    isMergeOrigin: hasVMerge && vMergeVal === 'restart',
                    isMergeContinuation: false,
                    getText: () => blocks.map(b => b.acceptedText).join('\n')
                };

                // Register as vMerge origin if applicable
                if (hasVMerge && vMergeVal === 'restart') {
                    for (let s = 0; s < colSpan; s++) {
                        vMergeOrigins.set(gridCol + s, { originRow: rowIdx, cell });
                    }
                } else {
                    // Clear vMerge origin for these columns if it was a restart or no vMerge
                    for (let s = 0; s < colSpan; s++) {
                        vMergeOrigins.delete(gridCol + s);
                    }
                }
            }

            // Place cell in grid (spanning multiple columns if needed)
            for (let spanOffset = 0; spanOffset < colSpan; spanOffset++) {
                const targetCol = gridCol + spanOffset;
                if (targetCol < colCount) {
                    grid[rowIdx][targetCol] = cell;
                    cellMap.set(`${rowIdx},${targetCol}`, cell);
                }
            }

            gridCol += colSpan;
        }
    }

    return {
        rowCount,
        colCount,
        grid,
        cellMap,
        tblPrXml: extractTblPr(tableNode),
        tblGridXml: extractTblGrid(tableNode),
        trPrList: Array.from(trElements).map(tr => extractTrPr(tr))
    };
}

function createRegularCell(rowIdx, gridCol, colSpan, tc, tcPr) {
    const blocks = parseCellBlocks(tc);
    return {
        gridRow: rowIdx,
        gridCol,
        rowSpan: 1,
        colSpan,
        tcNode: tc,
        blocks,
        tcPrXml: serializeTcPr(tcPr),
        isMergeOrigin: false,
        isMergeContinuation: false,
        getText: () => blocks.map(b => b.acceptedText).join('\n')
    };
}

function parseCellBlocks(tcNode) {
    const paragraphs = tcNode.getElementsByTagNameNS(NS_W, 'p').length > 0 ?
        tcNode.getElementsByTagNameNS(NS_W, 'p') :
        tcNode.getElementsByTagName('w:p');
    const cellBlocks = [];

    for (const p of paragraphs) {
        // We need a full paragraph OOXML string for ingestOoxml
        // For efficiency, we can wrap the single paragraph node
        const pXml = new XMLSerializer().serializeToString(p);
        const { runModel, acceptedText, pPr } = ingestOoxml(pXml);
        cellBlocks.push({
            runModel,
            acceptedText,
            pPr
        });
    }
    return cellBlocks;
}

function serializeTcPr(tcPrNode) {
    if (!tcPrNode) return '<w:tcPr/>';
    return new XMLSerializer().serializeToString(tcPrNode);
}

function extractTblPr(tableNode) {
    const tblPr = tableNode.getElementsByTagNameNS(NS_W, 'tblPr')[0] ||
        tableNode.getElementsByTagName('w:tblPr')[0];
    return tblPr ? new XMLSerializer().serializeToString(tblPr) : '<w:tblPr/>';
}

function extractTblGrid(tableNode) {
    const tblGrid = tableNode.getElementsByTagNameNS(NS_W, 'tblGrid')[0] ||
        tableNode.getElementsByTagName('w:tblGrid')[0];
    return tblGrid ? new XMLSerializer().serializeToString(tblGrid) : '<w:tblGrid/>';
}

function extractTrPr(trNode) {
    const trPr = trNode.getElementsByTagNameNS(NS_W, 'trPr')[0] ||
        trNode.getElementsByTagName('w:trPr')[0];
    return trPr ? new XMLSerializer().serializeToString(trPr) : '<w:trPr/>';
}
