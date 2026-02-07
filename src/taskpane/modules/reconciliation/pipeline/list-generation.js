/**
 * List generation flow extracted from ReconciliationPipeline.
 */

import { preprocessMarkdown } from './markdown-processor.js';
import { matchListMarker, extractListMarker, stripListMarker } from './list-markers.js';
import { serializeToOoxml } from './serialization.js';
import { generateTableOoxml } from '../services/table-reconciliation.js';
import { parseTable } from './content-analysis.js';
import { log } from '../adapters/logger.js';

/**
 * Executes list generation when paragraph content expands into list/table blocks.
 *
 * @param {Object} options - Generation options
 * @param {string} options.cleanText - Preprocessed target text
 * @param {Object|null} options.numberingContext - Existing numbering context
 * @param {Array} [options.originalRunModel=[]] - Original run model
 * @param {string} [options.originalText=''] - Original text fallback
 * @param {boolean} [options.generateRedlines=true] - Track-change toggle
 * @param {string} [options.author='AI'] - Author metadata
 * @param {string|null} [options.font=null] - Optional font
 * @param {import('../services/numbering-service.js').NumberingService} options.numberingService - Numbering service
 * @returns {Promise<import('../core/types.js').ReconciliationResult>}
 */
export async function executeListGeneration(options) {
    const {
        cleanText,
        numberingContext,
        originalRunModel = [],
        originalText = '',
        generateRedlines = true,
        author = 'AI',
        font = null,
        numberingService
    } = options;

    const lineMetadata = buildLineMetadata(cleanText);
    const rawLines = lineMetadata.map(line => line.raw);
    const results = [];

    let deletionRuns = [];
    if (generateRedlines) {
        if (originalRunModel && originalRunModel.length > 0) {
            deletionRuns = originalRunModel
                .filter(run => run.kind === 'text' || run.kind === 'run')
                .map(run => ({ ...run, kind: 'deletion', author }));
        } else if (originalText && originalText.trim().length > 0) {
            const trimmed = originalText.trim();
            deletionRuns = [{
                kind: 'deletion',
                text: trimmed,
                author,
                startOffset: 0,
                endOffset: trimmed.length
            }];
        }
    }

    const indentStep = detectIndentationStep(rawLines);
    log(`[ListGen] Detected indentation step: ${indentStep} spaces/chars`);

    const firstMarker = lineMetadata.find(line => line.marker)?.marker || '';
    const { format: defaultFormat } = numberingService.detectNumberingFormat(firstMarker);
    log(`[ListGen] Detected primary marker: "${firstMarker}", format: ${defaultFormat}`);

    for (let i = 0; i < lineMetadata.length; i++) {
        const tableBlock = collectMarkdownTableBlock(lineMetadata, i);
        if (tableBlock) {
            const tableData = parseTable(tableBlock.tableText);
            if (tableData.headers.length > 0 || tableData.rows.length > 0) {
                if (generateRedlines && results.length === 0 && deletionRuns.length > 0) {
                    results.push(serializeToOoxml(deletionRuns, null, [], {
                        author,
                        generateRedlines,
                        font
                    }));
                }

                results.push(generateTableOoxml(tableData, { generateRedlines, author }));
                i = tableBlock.endIndex;
                continue;
            }
        }

        const line = lineMetadata[i];
        const entry = buildListEntry(
            line,
            i,
            indentStep,
            numberingContext,
            numberingService,
            generateRedlines,
            author,
            font,
            deletionRuns
        );
        results.push(entry.ooxml);
    }

    const numberingXml = numberingService.generateNumberingXml();
    const finalOoxml = results.join('');
    const blankParagraph = '<w:p><w:pPr></w:pPr></w:p>';
    const oxmlWithSpacing = finalOoxml + blankParagraph;

    log(`[ListGen] ✅ Generated OOXML for ${results.length} list items, total length: ${oxmlWithSpacing.length}`);
    log(`[ListGen] First 200 chars: ${oxmlWithSpacing.substring(0, 200)}...`);

    return {
        ooxml: oxmlWithSpacing,
        isValid: true,
        warnings: ['Paragraph expanded to list fragment'],
        type: 'fragment',
        includeNumbering: true,
        numberingXml
    };
}

/**
 * Heuristically detects indentation step (spaces/tabs per level).
 *
 * @param {string[]} lines - Raw lines
 * @returns {number}
 */
export function detectIndentationStep(lines) {
    const indentations = lines
        .map(line => line.match(/^(\s*)/)[0].length)
        .filter(length => length > 0)
        .sort((a, b) => a - b);

    if (indentations.length === 0) return 2;

    let minJump = indentations[0];
    for (let i = 1; i < indentations.length; i++) {
        const jump = indentations[i] - indentations[i - 1];
        if (jump > 0 && jump < minJump) {
            minJump = jump;
        }
    }

    return minJump || 2;
}

function buildLineMetadata(cleanText) {
    return cleanText
        .split('\n')
        .filter(line => line.trim().length > 0)
        .map(raw => {
            const markerMatch = matchListMarker(raw);
            const headerMatch = raw.match(/^\s*(#{1,9})\s+(.*)/);
            return {
                raw,
                marker: markerMatch ? markerMatch[2].trim() : '',
                headerMatch,
                indentSize: (raw.match(/^(\s*)/)?.[1].length) || 0,
                isTableLine: /^\s*\|/.test(raw),
                isTableSeparator: /^\s*\|?[\s:-]*-[-\s|:]*\|?\s*$/.test(raw)
            };
        });
}

function collectMarkdownTableBlock(lineMetadata, index) {
    const current = lineMetadata[index];
    const next = lineMetadata[index + 1];

    if (!current?.isTableLine || !next?.isTableLine || !next?.isTableSeparator) {
        return null;
    }

    const tableLines = [];
    let cursor = index;
    while (cursor < lineMetadata.length && lineMetadata[cursor].isTableLine) {
        tableLines.push(lineMetadata[cursor].raw);
        cursor++;
    }

    return {
        tableText: tableLines.join('\n'),
        endIndex: cursor - 1
    };
}

function buildListEntry(
    line,
    lineIndex,
    indentStep,
    numberingContext,
    numberingService,
    generateRedlines,
    author,
    font,
    deletionRuns
) {
    let pPrXml = '';
    let segmentText = '';

    if (line.headerMatch) {
        const level = Math.min(line.headerMatch[1].length, 9);
        const outlineLevel = Math.min(level - 1, 8);
        const headingSizes = [32, 28, 26, 24, 22, 20, 20, 20, 20];
        const headingSize = headingSizes[level - 1] || headingSizes[headingSizes.length - 1];
        segmentText = line.headerMatch[2].trim();
        pPrXml = `<w:pPr><w:pStyle w:val="Heading${level}"/><w:outlineLvl w:val="${outlineLevel}"/><w:rPr><w:b/><w:sz w:val="${headingSize}"/><w:szCs w:val="${headingSize}"/></w:rPr></w:pPr>`;
    } else if (line.marker) {
        const lineFormat = numberingService.detectNumberingFormat(line.marker);
        const indentLevel = indentStep > 0 ? Math.floor(line.indentSize / indentStep) : 0;
        const contextLevel = numberingContext?.ilvl || 0;
        const ilvl = lineFormat.format === 'outline'
            ? Math.min(8, lineFormat.depth)
            : Math.min(8, indentLevel + contextLevel);

        segmentText = stripListMarker(line.raw);
        const numId = numberingService.getOrCreateNumId({ type: lineFormat.format }, numberingContext);
        pPrXml = numberingService.buildListPPr(numId, ilvl);
    } else {
        segmentText = line.raw;
    }

    const { cleanText, formatHints } = preprocessMarkdown(segmentText);
    const runModel = [];

    if (lineIndex === 0 && deletionRuns.length > 0) {
        runModel.push(...deletionRuns);
    }

    runModel.push({
        kind: generateRedlines ? 'insertion' : 'run',
        text: cleanText,
        author,
        startOffset: 0,
        endOffset: cleanText.length
    });

    return {
        ooxml: serializeToOoxml(runModel, pPrXml, formatHints, {
            author,
            generateRedlines,
            font
        })
    };
}
