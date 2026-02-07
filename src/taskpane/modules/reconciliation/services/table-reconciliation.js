/**
 * OOXML Reconciliation Pipeline - Table Reconciliation
 * 
 * Logic for reconciling tables using a Virtual Grid to handle merged cells.
 */

import { computeWordLevelDiffOps } from '../pipeline/diff-engine.js';
import { splitRunsAtDiffBoundaries, applyPatches } from '../pipeline/patching.js';
import { serializeToOoxml } from '../pipeline/serialization.js';
import { NS_W, getNextRevisionId, escapeXml, RunKind } from '../core/types.js';
import { preprocessMarkdown } from '../pipeline/markdown-processor.js';

/**
 * Generates a new w:tbl OOXML structure from Markdown table data.
 * This enables pure OOXML table creation without Word JS API.
 * 
 * @param {Object} tableData - Parsed Markdown table { headers, rows, hasHeader }
 * @param {Object} options - { generateRedlines, author }
 * @returns {string} Complete w:tbl OOXML
 */
export function generateTableOoxml(tableData, options = {}) {
    const { generateRedlines = false, author = 'AI' } = options;
    const date = new Date().toISOString();
    const revId = generateRedlines ? getNextRevisionId() : null;

    // Determine number of columns
    const numCols = tableData.headers?.length || (tableData.rows?.[0]?.length || 1);

    // Build default table properties (100% width, single borders)
    const tblPr = `
        <w:tblPr>
            <w:tblW w:w="5000" w:type="pct"/>
            <w:tblBorders>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            </w:tblBorders>
            <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
        </w:tblPr>
    `.trim();

    // Build grid columns (equal width)
    const gridCols = Array(numCols).fill('<w:gridCol/>').join('');
    const tblGrid = `<w:tblGrid>${gridCols}</w:tblGrid>`;

    // Build all rows
    const allRows = tableData.hasHeader ? [tableData.headers, ...tableData.rows] : tableData.rows;
    let rowsXml = '';

    for (let r = 0; r < allRows.length; r++) {
        const isHeaderRow = tableData.hasHeader && r === 0;
        const row = allRows[r] || [];
        let cellsXml = '';

        for (let c = 0; c < numCols; c++) {
            const cellText = row[c] || '';
            const { cleanText, formatHints } = preprocessMarkdown(cellText);

            // Build run model for the cell
            const runModel = [{
                kind: generateRedlines ? RunKind.INSERTION : RunKind.TEXT,
                text: cleanText,
                rPrXml: isHeaderRow ? '<w:rPr><w:b/></w:rPr>' : '',
                author,
                startOffset: 0,
                endOffset: cleanText.length
            }];

            const runsOoxml = serializeToOoxml(runModel, null, formatHints, { author, generateRedlines });

            // Cell properties
            const tcPr = '<w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>';
            // serializeToOoxml already returns one or more w:p blocks.
            cellsXml += `<w:tc>${tcPr}${runsOoxml}</w:tc>`;
        }

        // Row properties
        const trPr = '<w:trPr/>';
        rowsXml += `<w:tr>${trPr}${cellsXml}</w:tr>`;
    }

    // Build the table
    let tableXml = `<w:tbl>${tblPr}${tblGrid}${rowsXml}</w:tbl>`;

    // Wrap entire table in w:ins if generating redlines
    if (generateRedlines && revId) {
        tableXml = `<w:ins w:id="${revId}" w:author="${escapeXml(author)}" w:date="${date}">${tableXml}</w:ins>`;
    }

    return tableXml;
}

/**
 * Computes operations to reconcile two tables.
 * 
 * @param {Object} oldGrid - Original Virtual Grid model
 * @param {Object} newTableData - Parsed markdown table data
 * @returns {Array} List of operations
 */

export function diffTablesWithVirtualGrid(oldGrid, newTableData) {
    const operations = [];
    const { headers, rows: newRowsData, hasHeader } = newTableData;
    const allNewRows = hasHeader ? [headers, ...newRowsData] : newRowsData;

    const maxRows = Math.max(oldGrid.rowCount, allNewRows.length);

    for (let row = 0; row < maxRows; row++) {
        // Row insertion
        if (row >= oldGrid.rowCount && allNewRows[row]) {
            operations.push({
                type: 'row_insert',
                gridRow: row,
                cells: allNewRows[row]
            });
            continue;
        }

        // Row deletion
        if (row < oldGrid.rowCount && !allNewRows[row]) {
            operations.push({
                type: 'row_delete',
                gridRow: row
            });
            continue;
        }

        const newRow = allNewRows[row];

        for (let col = 0; col < oldGrid.colCount; col++) {
            const oldCell = oldGrid.grid[row]?.[col];
            const newCellText = newRow?.[col];

            if (!oldCell) continue;
            if (oldCell.isMergeContinuation) continue;
            if (col > oldCell.gridCol) continue;

            if (newCellText !== undefined) {
                const oldText = oldCell.getText();
                if (oldText !== newCellText) {
                    operations.push({
                        type: 'cell_modify',
                        gridRow: row,
                        gridCol: col,
                        originalCell: oldCell,
                        newText: newCellText
                    });
                }
            }
        }
    }

    return operations;
}

/**
 * Serializes the virtual grid back to OOXML after applying operations.
 * 
 * @param {Object} grid - The Virtual Grid
 * @param {Array} operations - Diff operations
 * @param {Object} options - Options (generateRedlines, author)
 * @returns {string} Reconciled w:tbl OOXML
 */
export function serializeVirtualGridToOoxml(grid, operations, options) {
    const { generateRedlines, author } = options;
    const date = new Date().toISOString();

    let rowsXml = '';

    for (let row = 0; row < grid.rowCount; row++) {
        const rowDeleteOp = operations.find(o => o.type === 'row_delete' && o.gridRow === row);

        if (rowDeleteOp && !generateRedlines) continue;

        let cellsXml = '';
        let col = 0;

        while (col < grid.colCount) {
            const cell = grid.grid[row][col];

            if (!cell) {
                col++;
                continue;
            }

            if (cell.isMergeContinuation) {
                col++;
                continue;
            }

            const modOp = operations.find(o =>
                o.type === 'cell_modify' &&
                o.gridRow === row &&
                o.gridCol === col
            );

            let cellContent;
            if (modOp) {
                cellContent = reconcileCellContent(cell, modOp.newText, options);
            } else {
                cellContent = serializeCellBlocks(cell.blocks);
            }

            cellsXml += buildTcXml(cell, cellContent, options);
            col += cell.colSpan;
        }

        let trPr = grid.trPrList[row] || '<w:trPr/>';

        if (rowDeleteOp && generateRedlines) {
            const revId = getNextRevisionId();
            const delMark = `<w:del w:id="${revId}" w:author="${escapeXml(author)}" w:date="${date}"/>`;
            if (trPr.includes('</w:trPr>')) {
                trPr = trPr.replace('</w:trPr>', `${delMark}</w:trPr>`);
            } else {
                trPr = `<w:trPr>${delMark}</w:trPr>`;
            }
        }

        rowsXml += `<w:tr>${trPr}${cellsXml}</w:tr>`;
    }

    // Handle row insertions
    const insertOps = operations.filter(o => o.type === 'row_insert').sort((a, b) => a.gridRow - b.gridRow);
    for (const op of insertOps) {
        let cellsXml = '';
        const revId = generateRedlines ? getNextRevisionId() : null;

        for (const cellText of op.cells) {
            const { cleanText, formatHints } = preprocessMarkdown(cellText);
            // Simple run model for new cell
            const runModel = [{
                kind: generateRedlines ? RunKind.INSERTION : RunKind.TEXT,
                text: cleanText,
                rPrXml: '',
                author,
                startOffset: 0,
                endOffset: cleanText.length
            }];

            const runsOoxml = serializeToOoxml(runModel, null, formatHints, { author, generateRedlines });
            // serializeToOoxml already returns one or more w:p blocks.
            cellsXml += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${runsOoxml}</w:tc>`;
        }

        let trPr = '<w:trPr/>';
        if (generateRedlines) {
            trPr = `<w:trPr><w:ins w:id="${revId}" w:author="${escapeXml(author)}" w:date="${date}"/></w:trPr>`;
        }

        rowsXml += `<w:tr>${trPr}${cellsXml}</w:tr>`;
    }

    return `
        <w:tbl>
            ${grid.tblPrXml}
            ${grid.tblGridXml}
            ${rowsXml}
        </w:tbl>
    `;
}

function reconcileCellContent(cell, newText, options) {
    const { generateRedlines, author } = options;
    const { cleanText, formatHints } = preprocessMarkdown(newText);

    // For now, satisfy with single block or join blocks
    // In a full implementation, we'd diff paragraphs within the cell
    const oldText = cell.getText();
    const diffOps = computeWordLevelDiffOps(oldText, cleanText);

    // Use the first block as reference for formatting if available
    const baseBlock = cell.blocks[0] || { runModel: [], pPr: null };
    const splitModel = splitRunsAtDiffBoundaries(baseBlock.runModel, diffOps);
    const patchedModel = applyPatches(splitModel, diffOps, {
        generateRedlines,
        author,
        formatHints
    });

    const runsOoxml = serializeToOoxml(patchedModel, baseBlock.pPr, formatHints, { author, generateRedlines });
    return runsOoxml;
}

function serializeCellBlocks(blocks) {
    return blocks.map(b => {
        const runsOoxml = serializeToOoxml(b.runModel, b.pPr, [], {});
        // serializeToOoxml already returns one or more w:p blocks.
        return runsOoxml;
    }).join('');
}

function buildTcXml(cell, content, options) {
    let tcPr = cell.tcPrXml;

    // Ensure gridSpan is preserved
    if (cell.colSpan > 1 && !tcPr.includes('gridSpan')) {
        tcPr = tcPr.replace('</w:tcPr>', `<w:gridSpan w:val="${cell.colSpan}"/></w:tcPr>`);
    } else if (cell.colSpan > 1 && tcPr === '<w:tcPr/>') {
        tcPr = `<w:tcPr><w:gridSpan w:val="${cell.colSpan}"/></w:tcPr>`;
    }

    // Ensure vMerge is preserved for origin cells
    if (cell.rowSpan > 1 && cell.isMergeOrigin && !tcPr.includes('vMerge')) {
        tcPr = tcPr.replace('</w:tcPr>', '<w:vMerge w:val="restart"/></w:tcPr>');
    } else if (cell.rowSpan > 1 && cell.isMergeOrigin && tcPr === '<w:tcPr/>') {
        tcPr = '<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>';
    }

    return `<w:tc>${tcPr}${content}</w:tc>`;
}
