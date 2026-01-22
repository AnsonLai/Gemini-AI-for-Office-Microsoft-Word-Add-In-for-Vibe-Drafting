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

        // Skip deleted content (w:del) - not part of accepted text
        if (nodeName === 'w:del') {
            // Store deletion for reconstruction if needed
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

        // Insertions: content IS part of accepted text (recurse into it)
        if (nodeName === 'w:ins') {
            const result = processNodeRecursive(child, localOffset, runModel);
            localOffset = result.offset;
            text += result.text;
            continue;
        }

        // Hyperlinks: preserve structure, extract text
        if (nodeName === 'w:hyperlink') {
            const hyperlinkEntry = processHyperlink(child, localOffset);
            runModel.push(hyperlinkEntry);
            localOffset += hyperlinkEntry.text.length;
            text += hyperlinkEntry.text;
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

        // SmartTag and SDT - stub for future container support
        if (nodeName === 'w:sdt' || nodeName === 'w:smartTag') {
            // For now, just recurse into content
            const result = processNodeRecursive(child, localOffset, runModel);
            localOffset = result.offset;
            text += result.text;
            continue;
        }
    }

    return { offset: localOffset, text };
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
 * Extracts text from any OOXML node (utility function).
 * 
 * @param {Element} node - The node to extract text from
 * @returns {string}
 */
export function extractTextFromNode(node) {
    let text = '';
    const textNodes = node.getElementsByTagNameNS(NS_W, 't');
    for (const t of textNodes) {
        text += t.textContent || '';
    }
    return text;
}
