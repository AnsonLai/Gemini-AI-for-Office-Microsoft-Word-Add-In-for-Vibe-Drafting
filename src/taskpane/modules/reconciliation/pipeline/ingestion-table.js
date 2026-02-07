/**
 * OOXML Reconciliation Pipeline - Table Ingestion
 *
 * Builds virtual-grid table models with merged-cell awareness.
 */

import { NS_W } from '../core/types.js';
import { serializeXml } from '../adapters/xml-adapter.js';
import {
    getElementsByTagNS,
    getElementsByTagNSOrTag,
    getFirstElementByTag,
    getFirstElementByTagNS
} from '../core/xml-query.js';
import { ingestParagraphElement } from './ingestion-paragraph.js';

/**
 * Ingests a table into a virtual grid model to handle merged cells.
 *
 * @param {Element} tableNode - w:tbl element
 * @returns {Object}
 */
export function ingestTableToVirtualGrid(tableNode) {
    const tblGrid = getFirstElementByTagNS(tableNode, NS_W, 'tblGrid');
    const gridCols = tblGrid ? getElementsByTagNS(tblGrid, NS_W, 'gridCol') : [];
    const colCount = gridCols.length;

    const trElements = getElementsByTagNS(tableNode, NS_W, 'tr');
    const rowCount = trElements.length;

    const grid = Array.from({ length: rowCount }, () =>
        Array.from({ length: colCount }, () => null)
    );
    const cellMap = new Map();
    const vMergeOrigins = new Map();

    for (let rowIdx = 0; rowIdx < trElements.length; rowIdx++) {
        const tr = trElements[rowIdx];
        const tcElements = getElementsByTagNS(tr, NS_W, 'tc');
        let gridCol = 0;

        for (let tcIdx = 0; tcIdx < tcElements.length; tcIdx++) {
            const tc = tcElements[tcIdx];
            const tcPr = getFirstElementByTagNS(tc, NS_W, 'tcPr');

            while (gridCol < colCount && grid[rowIdx][gridCol] !== null) {
                gridCol++;
            }
            if (gridCol >= colCount) break;

            const gridSpanEl = tcPr ? (getFirstElementByTagNS(tcPr, NS_W, 'gridSpan') || getFirstElementByTag(tcPr, 'w:gridSpan')) : null;
            const colSpan = parseInt(gridSpanEl?.getAttribute('w:val') || '1', 10);

            const vMergeEl = tcPr ? (getFirstElementByTagNS(tcPr, NS_W, 'vMerge') || getFirstElementByTag(tcPr, 'w:vMerge')) : null;
            const vMergeVal = vMergeEl?.getAttribute('w:val');
            const hasVMerge = vMergeEl !== null;

            let cell;
            if (hasVMerge && vMergeVal !== 'restart') {
                const origin = vMergeOrigins.get(gridCol);
                if (origin) {
                    origin.cell.rowSpan++;
                    cell = {
                        gridRow: rowIdx,
                        gridCol,
                        rowSpan: 0,
                        colSpan,
                        tcNode: tc,
                        blocks: [],
                        tcPrXml: serializeTcPr(tcPr),
                        isMergeOrigin: false,
                        isMergeContinuation: true,
                        mergeOrigin: origin.cell
                    };
                } else {
                    cell = createRegularCell(rowIdx, gridCol, colSpan, tc, tcPr);
                }
            } else {
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
                    getText: () => blocks.map(block => block.acceptedText).join('\n')
                };

                if (hasVMerge && vMergeVal === 'restart') {
                    for (let span = 0; span < colSpan; span++) {
                        vMergeOrigins.set(gridCol + span, { originRow: rowIdx, cell });
                    }
                } else {
                    for (let span = 0; span < colSpan; span++) {
                        vMergeOrigins.delete(gridCol + span);
                    }
                }
            }

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
        getText: () => blocks.map(block => block.acceptedText).join('\n')
    };
}

function parseCellBlocks(tcNode) {
    const paragraphs = getElementsByTagNSOrTag(tcNode, NS_W, 'p');
    return paragraphs.map(paragraph => {
        const { runModel, acceptedText, pPr } = ingestParagraphElement(paragraph);
        return { runModel, acceptedText, pPr };
    });
}

function serializeTcPr(tcPrNode) {
    if (!tcPrNode) return '<w:tcPr/>';
    return serializeXml(tcPrNode);
}

function extractTblPr(tableNode) {
    const tblPr = getFirstElementByTagNS(tableNode, NS_W, 'tblPr');
    return tblPr ? serializeXml(tblPr) : '<w:tblPr/>';
}

function extractTblGrid(tableNode) {
    const tblGrid = getFirstElementByTagNS(tableNode, NS_W, 'tblGrid');
    return tblGrid ? serializeXml(tblGrid) : '<w:tblGrid/>';
}

function extractTrPr(trNode) {
    const trPr = getFirstElementByTagNS(trNode, NS_W, 'trPr');
    return trPr ? serializeXml(trPr) : '<w:trPr/>';
}
