/**
 * OOXML Engine V5.1 - Hybrid Mode
 *
 * Router/orchestrator for formatting/text reconciliation modes.
 */

import { preprocessMarkdown } from '../pipeline/markdown-processor.js';
import { diffTablesWithVirtualGrid, serializeVirtualGridToOoxml, generateTableOoxml } from '../services/table-reconciliation.js';
import { parseTable, ReconciliationPipeline } from '../pipeline/pipeline.js';
import { ingestTableToVirtualGrid } from '../pipeline/ingestion.js';
import { wrapInDocumentFragment } from '../pipeline/serialization.js';
import { NS_W } from '../core/types.js';
import { createParser, createSerializer, parseXml, serializeXml } from '../adapters/xml-adapter.js';
import { log, error } from '../adapters/logger.js';
import { extractFormattingFromOoxml, getDocumentParagraphs } from './format-extraction.js';
import {
    applyFormatRemovalAsSurgicalReplacement,
    applyFormatOnlyChangesSurgical,
    buildParagraphInfos,
    findMatchingParagraphInfo,
    getContainingParagraph
} from './format-application.js';
import { detectTableCellContext, serializeParagraphOnly } from './table-cell-context.js';
import { applySurgicalMode } from './surgical-mode.js';
import { applyReconstructionMode } from './reconstruction-mode.js';

/**
 * Applies redline track changes to OOXML by modifying the DOM in-place.
 *
 * @param {string} oxml - Original OOXML string
 * @param {string} originalText - Original plain text
 * @param {string} modifiedText - New text (may contain markdown)
 * @param {Object} [options={}] - Options
 * @param {string} [options.author='AI'] - Author for track changes
 * @returns {Promise<{ oxml: string, hasChanges: boolean }>}
 */
export async function applyRedlineToOxml(oxml, originalText, modifiedText, options = {}) {
    const generateRedlines = options.generateRedlines ?? true;
    const author = options.author || 'Gemini AI';
    const parser = createParser();
    const serializer = createSerializer();

    let xmlDoc;
    try {
        xmlDoc = parser.parseFromString(oxml, 'text/xml');
    } catch (e) {
        error('[OxmlEngine] Failed to parse OXML:', e);
        return { oxml, hasChanges: false };
    }

    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        error('[OxmlEngine] XML parse error:', parseError.textContent);
        return { oxml, hasChanges: false };
    }

    const initialTableCellContext = detectTableCellContext(xmlDoc, originalText);
    if (initialTableCellContext.hasTableWrapper && initialTableCellContext.targetParagraph && !options._isolatedTableCell) {
        log('[OxmlEngine] Isolating table-cell paragraph before diff');
        const isolatedOxml = serializeParagraphOnly(xmlDoc, initialTableCellContext.targetParagraph, serializer);
        return applyRedlineToOxml(isolatedOxml, originalText, modifiedText, {
            ...options,
            _isolatedTableCell: true
        });
    }

    const sanitizedText = sanitizeAiResponse(modifiedText);
    const { cleanText: cleanModifiedText, formatHints } = preprocessMarkdown(sanitizedText);

    const hasTextChanges = cleanModifiedText.trim() !== originalText.trim();
    const hasFormatHints = formatHints.length > 0;

    const { existingFormatHints, textSpans } = extractFormattingFromOoxml(xmlDoc);
    const hasExistingFormatting = existingFormatHints.length > 0;

    log(`[OxmlEngine] Text changes: ${hasTextChanges}, New format hints: ${formatHints.length}, Existing format hints: ${existingFormatHints.length}`);

    const needsFormatRemoval = !hasTextChanges && !hasFormatHints && hasExistingFormatting;

    if (!hasTextChanges && !hasFormatHints && !hasExistingFormatting) {
        log('[OxmlEngine] No text changes, no format hints, and no existing formatting detected');
        return { oxml, hasChanges: false };
    }

    if (needsFormatRemoval) {
        log('[OxmlEngine] Format REMOVAL detected: applying surgical replacement in OOXML');

        const tableCellCtx = detectTableCellContext(xmlDoc, originalText);
        let targetParagraph = tableCellCtx.targetParagraph || null;

        if (!targetParagraph) {
            const allParagraphs = getDocumentParagraphs(xmlDoc);
            const paragraphInfos = buildParagraphInfos(xmlDoc, allParagraphs, textSpans);
            const matchedInfo = findMatchingParagraphInfo(paragraphInfos, originalText);
            if (matchedInfo) {
                targetParagraph = matchedInfo.paragraph;
            }
        }

        let filteredHints = existingFormatHints;
        if (targetParagraph) {
            filteredHints = existingFormatHints.filter(hint => {
                const hintParagraph = getContainingParagraph(hint.run);
                return hintParagraph === targetParagraph;
            });
        }

        const removalResult = applyFormatRemovalAsSurgicalReplacement(
            xmlDoc,
            textSpans,
            filteredHints,
            serializer,
            author,
            generateRedlines
        );

        if (tableCellCtx.hasTableWrapper && targetParagraph) {
            return {
                oxml: serializeParagraphOnly(xmlDoc, targetParagraph, serializer),
                hasChanges: removalResult.hasChanges
            };
        }

        return removalResult;
    }

    if (!hasTextChanges && hasFormatHints) {
        log(`[OxmlEngine] Format-only change detected: ${formatHints.length} format hints`);

        const tableCellCtx = detectTableCellContext(xmlDoc, originalText);
        if (tableCellCtx.hasTableWrapper && tableCellCtx.targetParagraph) {
            log('[OxmlEngine] Table cell context: applying formatting to target paragraph only');

            const formatResult = applyFormatOnlyChangesSurgical(
                xmlDoc,
                originalText,
                formatHints,
                serializer,
                author,
                generateRedlines,
                textSpans
            );

            if (formatResult.useNativeApi) {
                return formatResult;
            }

            log('[OxmlEngine] Stripping table wrapper for table cell paragraph (format-only)');
            return {
                oxml: serializeParagraphOnly(xmlDoc, tableCellCtx.targetParagraph, serializer),
                hasChanges: formatResult.hasChanges
            };
        }

        return applyFormatOnlyChangesSurgical(xmlDoc, originalText, formatHints, serializer, author, generateRedlines, textSpans);
    }

    const tables = xmlDoc.getElementsByTagName('w:tbl');
    const hasTables = tables.length > 0;
    const isMarkdownTable = /^\|.+\|/.test(cleanModifiedText.trim()) && cleanModifiedText.includes('\n');
    const markersRegex = /^(\s*)((?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|[-*•])\s*)/m;
    const isTargetList = cleanModifiedText.includes('\n') && markersRegex.test(cleanModifiedText.trim());
    const tableCellContext = initialTableCellContext;

    log(`[OxmlEngine] Mode: ${hasTables ? 'SURGICAL' : 'RECONSTRUCTION'}, formatHints: ${formatHints.length}, isMarkdownTable: ${isMarkdownTable}, isTargetList: ${isTargetList}, isTableCellParagraph: ${tableCellContext.isTableCellParagraph}`);

    if (isMarkdownTable && !hasTables) {
        log('[OxmlEngine] Text-to-table transformation: generating new table from Markdown');
        return applyTextToTableTransformation(xmlDoc, cleanModifiedText, serializer, author, generateRedlines);
    }

    if (hasTables && isMarkdownTable) {
        return applyTableReconciliation(xmlDoc, cleanModifiedText, serializer, author, formatHints, generateRedlines);
    }
    if (hasTables) {
        const surgicalTarget = tableCellContext.hasTableWrapper && tableCellContext.targetParagraph
            ? tableCellContext.targetParagraph
            : null;
        if (surgicalTarget) {
            log('[OxmlEngine] Table cell edit: scoping surgical mode to target paragraph');
        }

        const result = applySurgicalMode(
            xmlDoc,
            originalText,
            cleanModifiedText,
            serializer,
            author,
            formatHints,
            generateRedlines,
            surgicalTarget
        );

        if (tableCellContext.hasTableWrapper && result.hasChanges && tableCellContext.targetParagraph) {
            log('[OxmlEngine] Stripping table wrapper for table cell paragraph (surgical mode)');
            return { oxml: serializeParagraphOnly(xmlDoc, tableCellContext.targetParagraph, serializer), hasChanges: true };
        }
        return result;
    }
    if (isTargetList) {
        log('[OxmlEngine] 🎯 Using reconciliation pipeline for list generation');
        const pipeline = new ReconciliationPipeline({ author, generateRedlines });
        const result = await pipeline.execute(oxml, modifiedText);

        if (result.isValid && result.ooxml && result.ooxml !== oxml) {
            log(`[OxmlEngine] Wrapping list OOXML with numbering definitions, includeNumbering=${result.includeNumbering}`);
            const wrapped = wrapInDocumentFragment(result.ooxml, {
                includeNumbering: result.includeNumbering || true,
                numberingXml: result.numberingXml
            });
            log(`[OxmlEngine] ✅ Wrapped OOXML length: ${wrapped.length}`);
            return { oxml: wrapped, hasChanges: true };
        }
        return { oxml, hasChanges: false };
    }

    return applyReconstructionMode(xmlDoc, originalText, cleanModifiedText, serializer, author, formatHints, generateRedlines);
}

/**
 * Applies structural reconciliation to tables using Virtual Grid.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} modifiedText - Markdown/target table text
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Author name
 * @param {Array} formatHints - Format hints (reserved)
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
function applyTableReconciliation(xmlDoc, modifiedText, serializer, author, formatHints, generateRedlines = true) {
    const tableNodes = Array.from(xmlDoc.getElementsByTagName('w:tbl'));
    const newTableData = parseTable(modifiedText);
    const hasNewContent = newTableData.rows.length > 0 || newTableData.headers.length > 0;

    if (tableNodes.length === 0 || !hasNewContent) {
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    const targetTable = tableNodes[0];
    const oldGrid = ingestTableToVirtualGrid(targetTable);
    const operations = diffTablesWithVirtualGrid(oldGrid, newTableData);

    if (operations.length === 0) {
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    const options = { generateRedlines, author };
    const reconciledOxml = serializeVirtualGridToOoxml(oldGrid, operations, options);

    const parser = createParser();
    const wrappedOxml = `<root xmlns:w="${NS_W}">${reconciledOxml}</root>`;
    const reconciledDoc = parser.parseFromString(wrappedOxml, 'application/xml');

    const parseError = reconciledDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        error('[OxmlEngine] Failed to parse reconciled table OOXML:', parseError.textContent);
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    const newTableNode = reconciledDoc.getElementsByTagName('w:tbl')[0];
    if (!newTableNode) {
        error('[OxmlEngine] No table found in reconciled OOXML');
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    const importedTable = xmlDoc.importNode(newTableNode, true);
    targetTable.parentNode.replaceChild(importedTable, targetTable);

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

/**
 * Transforms paragraph content into a new table from Markdown text.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} modifiedText - Markdown table text
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Author name
 * @param {boolean} generateRedlines - Track change toggle
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
function applyTextToTableTransformation(xmlDoc, modifiedText, serializer, author, generateRedlines) {
    const tableData = parseTable(modifiedText);

    if (!tableData || (tableData.rows.length === 0 && tableData.headers.length === 0)) {
        log('[OxmlEngine] Failed to parse table data from Markdown');
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    const tableOoxml = generateTableOoxml(tableData, { generateRedlines, author });

    const parser = createParser();
    const tableDoc = parser.parseFromString(
        `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${tableOoxml}</root>`,
        'application/xml'
    );

    const parseError = tableDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        error('[OxmlEngine] Failed to parse generated table OOXML:', parseError.textContent);
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    let newTableElement = tableDoc.getElementsByTagNameNS(NS_W, 'tbl')[0];
    if (!newTableElement) {
        newTableElement = tableDoc.getElementsByTagNameNS(NS_W, 'ins')[0];
    }
    if (!newTableElement) {
        error('[OxmlEngine] No table element found in generated OOXML');
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    let workingDoc = xmlDoc;
    let paragraphs = Array.from(workingDoc.getElementsByTagNameNS(NS_W, 'p'));

    if (paragraphs.length === 0) {
        log('[OxmlEngine] No paragraphs found to replace');
        return { oxml: serializer.serializeToString(workingDoc), hasChanges: false };
    }

    let firstParagraph = paragraphs[0];
    let parent = firstParagraph.parentNode;

    // If the input is a standalone paragraph document (root is w:p),
    // create a temporary w:document/w:body wrapper so table + paragraph siblings are valid XML.
    if (parent && parent.nodeType === 9) {
        const wrappedDoc = createParser().parseFromString(
            `<w:document xmlns:w="${NS_W}"><w:body/></w:document>`,
            'application/xml'
        );
        const wrappedBody = wrappedDoc.getElementsByTagNameNS(NS_W, 'body')[0];
        paragraphs.forEach(p => wrappedBody.appendChild(wrappedDoc.importNode(p, true)));

        workingDoc = wrappedDoc;
        paragraphs = Array.from(workingDoc.getElementsByTagNameNS(NS_W, 'p'));
        firstParagraph = paragraphs[0];
        parent = firstParagraph.parentNode;
    }

    const importedTable = workingDoc.importNode(newTableElement, true);

    if (generateRedlines) {
        const date = new Date().toISOString();
        paragraphs.forEach(p => {
            const runs = Array.from(p.getElementsByTagNameNS(NS_W, 'r'));
            runs.forEach(run => {
                const textNodes = Array.from(run.getElementsByTagNameNS(NS_W, 't'));
                textNodes.forEach(t => {
                    const text = t.textContent || '';
                    if (text.trim()) {
                        const delText = workingDoc.createElementNS(NS_W, 'w:delText');
                        delText.textContent = text;
                        t.parentNode.replaceChild(delText, t);
                    }
                });

                const del = workingDoc.createElementNS(NS_W, 'w:del');
                del.setAttribute('w:id', String(Math.floor(Math.random() * 100000)));
                del.setAttribute('w:author', author);
                del.setAttribute('w:date', date);
                run.parentNode.insertBefore(del, run);
                del.appendChild(run);
            });
        });
    } else {
        paragraphs.slice(1).forEach(p => p.parentNode.removeChild(p));
    }

    parent.insertBefore(importedTable, firstParagraph);

    if (!generateRedlines) {
        parent.removeChild(firstParagraph);
    }

    log('[OxmlEngine] Text-to-table transformation complete');
    return { oxml: serializer.serializeToString(workingDoc), hasChanges: true };
}

/**
 * Sanitizes AI response text by removing common prefixes.
 *
 * @param {string} text - AI response text
 * @returns {string}
 */
export function sanitizeAiResponse(text) {
    let cleaned = text;
    cleaned = cleaned.replace(/^(Here is the redline:|Here is the text:|Sure, I can help:|Here's the updated text:)\s*/i, '');
    cleaned = cleaned.replace(/\$\\text\{/g, '').replace(/\}\$/g, '');
    cleaned = cleaned.replace(/\$([^0-9\n]+?)\$/g, '$1');
    cleaned = cleaned.replace(/\\r\\n/g, '\n').replace(/\\n/g, '\n');
    return cleaned;
}

/**
 * Parses OOXML into a DOM document.
 *
 * @param {string} ooxmlString - OOXML text
 * @returns {Document}
 */
export function parseOoxml(ooxmlString) {
    return parseXml(ooxmlString, 'application/xml');
}

/**
 * Serializes a DOM document to OOXML text.
 *
 * @param {Node} doc - XML document/node
 * @returns {string}
 */
export function serializeOoxml(doc) {
    return serializeXml(doc);
}
