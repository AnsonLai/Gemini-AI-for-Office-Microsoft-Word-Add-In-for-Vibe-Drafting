/**
 * OOXML Reconciliation Pipeline - Main Pipeline
 * 
 * Orchestrates the reconciliation process from OOXML input to output.
 */

import { ingestOoxml } from './ingestion.js';
import { preprocessMarkdown } from './markdown-processor.js';
import { computeWordLevelDiffOps } from './diff-engine.js';
import { splitRunsAtDiffBoundaries, applyPatches } from './patching.js';
import { serializeToOoxml, wrapInDocumentFragment } from './serialization.js';
import { ContentType } from './types.js';
import { NumberingService } from './numbering-service.js';
import { detectNumberingContext } from './ingestion.js';
import { generateTableOoxml } from './table-reconciliation.js';

/**
 * Main reconciliation pipeline class.
 * Orchestrates the process of diffing and patching OOXML content.
 */
export class ReconciliationPipeline {
    /**
     * @param {Object} options - Pipeline options
     * @param {boolean} [options.generateRedlines=true] - Generate track changes
     * @param {string} [options.author='AI'] - Author for track changes
     * @param {boolean} [options.validateOutput=true] - Validate output before returning
     */
    constructor(options = {}) {
        this.generateRedlines = options.generateRedlines ?? true;
        this.author = options.author ?? 'AI';
        this.validateOutput = options.validateOutput ?? true;
        this.numberingService = options.numberingService || new NumberingService();
    }

    /**
     * Executes the reconciliation pipeline.
     * 
     * @param {string} originalOoxml - Original OOXML paragraph content
     * @param {string} newText - New text with optional markdown formatting
     * @returns {Promise<import('./types.js').ReconciliationResult>}
     */
    async execute(originalOoxml, newText) {
        const warnings = [];

        try {
            // Stage 1: Ingest OOXML
            const parser = new DOMParser();
            const doc = parser.parseFromString(originalOoxml, 'application/xml');
            const pElement = doc.getElementsByTagNameNS('*', 'p')[0];

            const { runModel, acceptedText, pPr } = ingestOoxml(originalOoxml);
            const numberingContext = pElement ? detectNumberingContext(pElement) : null;

            console.log(`[Reconcile] Ingested ${runModel.length} runs, ${acceptedText.length} chars, numbering:`, numberingContext);

            // Stage 2: Preprocess markdown
            const { cleanText, formatHints } = preprocessMarkdown(newText);
            console.log(`[Reconcile] Preprocessed: ${formatHints.length} format hints`);

            // Early exit if no change
            if (acceptedText === cleanText && formatHints.length === 0) {
                console.log('[Reconcile] No changes detected');
                return {
                    ooxml: originalOoxml,
                    isValid: true,
                    warnings: ['No changes detected']
                };
            }

            // Stage 3: Compute word-level diff
            const diffOps = computeWordLevelDiffOps(acceptedText, cleanText);

            // Detect if this is a list transformation (e.g., paragraph with newlines)
            const isTargetList = cleanText.includes('\n') && /^([-*+]|\d+\.)/m.test(cleanText);
            // Key fix: Only use executeListGeneration for EXPANSION (single paragraph -> list)
            // If acceptedText already has newlines, we ingested multiple paragraphs and should use surgical diff
            const isExpansion = isTargetList && !acceptedText.includes('\n');

            if (isExpansion) {
                console.log('[Reconcile] Detected list expansion from single paragraph');
                return this.executeListGeneration(cleanText, numberingContext, runModel);
            }

            console.log(`[Reconcile] Computed ${diffOps.length} diff operations`);

            // Stage 4: Pre-split runs at boundaries
            const splitModel = splitRunsAtDiffBoundaries(runModel, diffOps);
            console.log(`[Reconcile] Split into ${splitModel.length} runs`);

            // Stage 5: Apply patches
            const patchedModel = applyPatches(splitModel, diffOps, {
                generateRedlines: this.generateRedlines,
                author: this.author,
                formatHints
            });
            console.log(`[Reconcile] Patched model has ${patchedModel.length} runs`);

            // Stage 6: Serialize to OOXML
            const resultOoxml = serializeToOoxml(patchedModel, pPr, formatHints, {
                author: this.author
            });

            // Stage 7: Basic validation
            if (this.validateOutput) {
                const validation = this.validateBasic(resultOoxml);
                if (!validation.isValid) {
                    warnings.push(...validation.errors);
                }
            }

            return {
                ooxml: resultOoxml,
                isValid: warnings.length === 0,
                warnings
            };

        } catch (error) {
            console.error('[Reconcile] Pipeline error:', error);
            return {
                ooxml: originalOoxml,
                isValid: false,
                warnings: [`Pipeline error: ${error.message}`]
            };
        }
    }

    /**
     * Performs basic validation on generated OOXML.
     * 
     * @param {string} ooxml - Generated OOXML
     * @returns {{ isValid: boolean, errors: string[] }}
     */
    validateBasic(ooxml) {
        const errors = [];

        try {
            // Check for well-formed XML by wrapping in namespace container
            const wrappedXml = `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${ooxml}</root>`;
            const parser = new DOMParser();
            const doc = parser.parseFromString(wrappedXml, 'application/xml');

            const parseError = doc.getElementsByTagName('parsererror')[0];
            if (parseError) {
                errors.push('Generated OOXML is not well-formed XML: ' + parseError.textContent.substring(0, 100));
            }

            // Check for basic structure
            if (!ooxml.includes('<w:p')) {
                errors.push('Generated OOXML missing paragraph element');
            }

        } catch (e) {
            errors.push(`Validation error: ${e.message}`);
        }

        return {
            isValid: errors.length === 0,
            errors
        };
    }

    /**
     * Wraps the reconciled content for document insertion.
     * 
     * @param {string} ooxml - Reconciled OOXML paragraph
     * @param {Object} [options={}] - Options
     * @param {boolean} [options.includeNumbering=false] - Include numbering definitions
     * @returns {string} Wrapped document fragment
     */
    wrapForInsertion(ooxml, options = {}) {
        return wrapInDocumentFragment(ooxml, options);
    }

    /**
     * Executes list generation when a single paragraph expands into a list.
     * 
     * @param {string} cleanText - Preprocessed new text (markdown list)
     * @param {Object} numberingContext - Original numbering context
     * @param {Array} originalRunModel - Run model of the original paragraph (optional)
     * @param {string} originalText - Plain text of the original paragraph (optional, used if runModel not provided)
     */
    async executeListGeneration(cleanText, numberingContext, originalRunModel, originalText = '') {
        const lines = cleanText.split('\n').filter(l => l.trim().length > 0);
        const results = [];

        // Identify the original text to be deleted across the new paragraphs
        // For simplicity in list expansion, we'll put the deletion in the first generated paragraph
        let deletionRuns = [];
        if (this.generateRedlines) {
            if (originalRunModel && originalRunModel.length > 0) {
                // Use run model if provided
                deletionRuns = originalRunModel
                    .filter(r => r.kind === 'text' || r.kind === 'run')
                    .map(r => ({
                        ...r,
                        kind: 'deletion',
                        author: this.author
                    }));
            } else if (originalText && originalText.trim().length > 0) {
                // Create simple deletion run from text
                deletionRuns = [{
                    kind: 'deletion',
                    text: originalText.trim(),
                    author: this.author,
                    startOffset: 0,
                    endOffset: originalText.trim().length
                }];
            }
        }

        // Determine the primary list type and format from the first item
        let firstMarker = '';
        for (const line of lines) {
            const match = line.match(/^\s*((?:\d+(?:\.\d+)*\.?|\([a-zA-Z0-9ivxlc]+\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|-|\*|•)\s*)/);
            if (match) {
                firstMarker = match[1].trim();
                break;
            }
        }

        const { format } = this.numberingService.detectNumberingFormat(firstMarker);
        console.log(`[ListGen] Detected marker: "${firstMarker}", format: ${format}`);

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const indentMatch = line.match(/^(\s*)/);
            const indent = indentMatch ? indentMatch[1].length : 0;

            // 2 spaces = 1 level
            const ilvl = Math.min(8, Math.floor(indent / 2) + (numberingContext?.ilvl || 0));

            // Extract the marker for THIS line to see if it overrides the primary format 
            // (e.g. mixed lists, though we try to stay consistent)
            const markerMatch = line.match(/^\s*((?:\d+(?:\.\d+)*\.?|\([a-zA-Z0-9ivxlc]+\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|-|\*|•)\s*)/);
            const currentMarker = markerMatch ? markerMatch[1].trim() : '';
            const lineFormat = currentMarker ? this.numberingService.detectNumberingFormat(currentMarker).format : format;

            // Strip list markers from the text
            const rawText = line.trim().replace(/^((?:\d+(?:\.\d+)*\.?|\([a-zA-Z0-9ivxlc]+\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|-|\*|•)\s*)/, '');

            // Process markdown formatting (e.g., **bold**, *italic*)
            const { cleanText, formatHints } = preprocessMarkdown(rawText);

            const numId = this.numberingService.getOrCreateNumId({ type: lineFormat }, numberingContext);
            const pPrXml = this.numberingService.buildListPPr(numId, ilvl);

            const runModel = [];

            // Add deletion to the first paragraph only
            if (i === 0 && deletionRuns.length > 0) {
                runModel.push(...deletionRuns);
            }

            // Add the new text as an insertion (with clean text, formatting comes from hints)
            runModel.push({
                kind: this.generateRedlines ? 'insertion' : 'run',
                text: cleanText,
                author: this.author,
                startOffset: 0,
                endOffset: cleanText.length
            });

            // Pass formatHints to serialization for proper bold/italic/etc formatting
            const itemOoxml = serializeToOoxml(runModel, pPrXml, formatHints, { author: this.author });
            results.push(itemOoxml);
        }

        const numberingXml = this.numberingService.generateNumberingXml();

        return {
            ooxml: results.join(''),
            isValid: true,
            warnings: ['Paragraph expanded to list fragment'],
            type: 'fragment',
            includeNumbering: true,
            numberingXml: numberingXml
        };
    }

    /**
     * Executes table generation from markdown text.
     * 
     * @param {string} markdownTable - Markdown table text
     * @returns {Object} ReconciliationResult containing the table OOXML
     */
    executeTableGeneration(markdownTable) {
        const tableData = parseTable(markdownTable);
        if (tableData.rows.length === 0 && tableData.headers.length === 0) {
            return {
                ooxml: '',
                isValid: false,
                warnings: ['Could not parse Markdown table']
            };
        }

        const tableOoxml = generateTableOoxml(tableData, {
            generateRedlines: this.generateRedlines,
            author: this.author
        });

        return {
            ooxml: tableOoxml,
            isValid: true,
            warnings: [],
            includeNumbering: false
        };
    }
}

// ============================================================================
// STUBS FOR FUTURE FEATURES
// ============================================================================

/**
 * Detects content type (paragraph, list, table) from text.
 * STUB: Currently always returns PARAGRAPH.
 * 
 * @param {string} text - Text to analyze
 * @returns {ContentType}
 */
export function detectContentType(text) {
    // TODO: Implement content type detection
    console.log('[Stub] detectContentType - returning PARAGRAPH');
    return ContentType.PARAGRAPH;
}

/**
 * Parses list items from markdown-style list text.
 * STUB: Not yet implemented.
 * 
 * @param {string} text - List text
 * @returns {Array}
 */
export function parseListItems(text) {
    // TODO: Implement list parsing
    console.log('[Stub] parseListItems - not implemented');
    return [];
}

/**
 * Parses table from markdown-style table text.
 * 
 * @param {string} text - Table text
 * @returns {Object}
 */
export function parseTable(text) {
    const lines = text.split('\n').filter(l => l.trim().startsWith('|'));

    if (lines.length === 0) {
        return { headers: [], rows: [], hasHeader: false };
    }

    // Skip separator row (|---|---|)
    const dataLines = lines.filter(l => !l.includes('---'));

    const rows = dataLines.map(line => {
        return line
            .split('|')
            .slice(1, -1)  // Remove empty first/last from split
            .map(cell => cell.trim());
    });

    return {
        headers: rows[0] || [],
        rows: rows.slice(1),
        hasHeader: lines.some(l => l.includes('---'))
    };
}
