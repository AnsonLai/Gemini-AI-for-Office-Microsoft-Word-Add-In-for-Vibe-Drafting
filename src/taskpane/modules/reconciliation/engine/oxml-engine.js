/**
 * OOXML Engine V5.1 - Hybrid Mode
 *
 * Router/orchestrator for formatting/text reconciliation modes.
 */

import { preprocessMarkdown } from '../pipeline/markdown-processor.js';
import { isListTargetLoose } from '../pipeline/list-markers.js';
import { ReconciliationPipeline } from '../pipeline/pipeline.js';
import { wrapInDocumentFragment } from '../pipeline/serialization.js';
import {
    getElementsByTag,
    getXmlParseError
} from '../core/xml-query.js';
import { createParser, createSerializer, parseXml, serializeXml } from '../adapters/xml-adapter.js';
import { log, error } from '../adapters/logger.js';
import { extractFormattingFromOoxml } from './format-extraction.js';
import {
    applyFormatRemovalAsSurgicalReplacement,
    applyFormatOnlyChangesSurgical
} from './format-application.js';
import { buildParagraphInfos, findMatchingParagraphInfo, getContainingParagraph } from './format-paragraph-targeting.js';
import { detectTableCellContext, serializeParagraphOnly } from './table-cell-context.js';
import { applySurgicalMode } from './surgical-mode.js';
import { applyReconstructionMode } from './reconstruction-mode.js';
import { applyTableReconciliation, applyTextToTableTransformation } from './table-mode.js';

/**
 * Applies redline track changes to OOXML by modifying the DOM in-place.
 *
 * @param {string} oxml - Original OOXML string
 * @param {string} originalText - Original plain text
 * @param {string} modifiedText - New text (may contain markdown)
 * @param {Object} [options={}] - Options
 * @param {string} [options.author='AI'] - Author for track changes
 * @param {string|null} [options.targetParagraphId=null] - Preferred paragraph identity for table wrappers
 * @returns {Promise<{ oxml: string, hasChanges: boolean }>}
 */
export async function applyRedlineToOxml(oxml, originalText, modifiedText, options = {}) {
    const generateRedlines = options.generateRedlines ?? true;
    const author = options.author || 'Gemini AI';
    const parser = createParser();
    const serializer = createSerializer();
    const noChanges = () => ({ oxml, hasChanges: false });

    let xmlDoc;
    try {
        xmlDoc = parser.parseFromString(oxml, 'text/xml');
    } catch (e) {
        error('[OxmlEngine] Failed to parse OXML:', e);
        return noChanges();
    }

    const parseError = getXmlParseError(xmlDoc);
    if (parseError) {
        error('[OxmlEngine] XML parse error:', parseError.textContent);
        return noChanges();
    }

    const initialTableCellContext = detectTableCellContext(xmlDoc, originalText, options);
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

    const { existingFormatHints, textSpans, paragraphs } = extractFormattingFromOoxml(xmlDoc);
    const hasExistingFormatting = existingFormatHints.length > 0;
    let paragraphInfos = null;
    const getParagraphInfos = () => {
        if (!paragraphInfos) {
            paragraphInfos = buildParagraphInfos(xmlDoc, paragraphs, textSpans);
        }
        return paragraphInfos;
    };

    log(`[OxmlEngine] Text changes: ${hasTextChanges}, New format hints: ${formatHints.length}, Existing format hints: ${existingFormatHints.length}`);

    const needsFormatRemoval = !hasTextChanges && !hasFormatHints && hasExistingFormatting;

    if (!hasTextChanges && !hasFormatHints && !hasExistingFormatting) {
        log('[OxmlEngine] No text changes, no format hints, and no existing formatting detected');
        return noChanges();
    }

    if (needsFormatRemoval) {
        log('[OxmlEngine] Format REMOVAL detected: applying surgical replacement in OOXML');

        const tableCellCtx = initialTableCellContext;
        let targetParagraph = tableCellCtx.targetParagraph || null;

        if (!targetParagraph) {
            const matchedInfo = findMatchingParagraphInfo(getParagraphInfos(), originalText);
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

        const tableCellCtx = initialTableCellContext;
        const precomputedFormatContext = {
            textSpans,
            paragraphs,
            paragraphInfos: getParagraphInfos()
        };
        if (tableCellCtx.hasTableWrapper && tableCellCtx.targetParagraph) {
            log('[OxmlEngine] Table cell context: applying formatting to target paragraph only');

            const formatResult = applyFormatOnlyChangesSurgical(
                xmlDoc,
                originalText,
                formatHints,
                serializer,
                author,
                generateRedlines,
                precomputedFormatContext
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

        return applyFormatOnlyChangesSurgical(xmlDoc, originalText, formatHints, serializer, author, generateRedlines, precomputedFormatContext);
    }

    const tables = getElementsByTag(xmlDoc, 'w:tbl');
    const hasTables = tables.length > 0;
    const isMarkdownTable = /^\|.+\|/.test(cleanModifiedText.trim()) && cleanModifiedText.includes('\n');
    const isTargetList = isListTargetLoose(cleanModifiedText);
    const tableCellContext = initialTableCellContext;

    log(`[OxmlEngine] Mode: ${hasTables ? 'SURGICAL' : 'RECONSTRUCTION'}, formatHints: ${formatHints.length}, isMarkdownTable: ${isMarkdownTable}, isTargetList: ${isTargetList}, isTableCellParagraph: ${tableCellContext.isTableCellParagraph}`);

    if (isMarkdownTable && !hasTables) {
        log('[OxmlEngine] Text-to-table transformation: generating new table from Markdown');
        return applyTextToTableTransformation(xmlDoc, cleanModifiedText, serializer, parser, author, generateRedlines);
    }

    if (hasTables && isMarkdownTable) {
        return applyTableReconciliation(xmlDoc, cleanModifiedText, serializer, parser, author, generateRedlines);
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
                includeNumbering: result.includeNumbering ?? true,
                numberingXml: result.numberingXml
            });
            log(`[OxmlEngine] ✅ Wrapped OOXML length: ${wrapped.length}`);
            return { oxml: wrapped, hasChanges: true };
        }
        return noChanges();
    }

    return applyReconstructionMode(xmlDoc, originalText, cleanModifiedText, serializer, author, formatHints, generateRedlines);
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
