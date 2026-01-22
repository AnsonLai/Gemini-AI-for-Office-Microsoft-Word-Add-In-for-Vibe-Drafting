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
            const { runModel, acceptedText, pPr } = ingestOoxml(originalOoxml);
            console.log(`[Reconcile] Ingested ${runModel.length} runs, ${acceptedText.length} chars`);

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
     * @returns {string} Wrapped document fragment
     */
    wrapForInsertion(ooxml) {
        return wrapInDocumentFragment(ooxml);
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
 * STUB: Not yet implemented.
 * 
 * @param {string} text - Table text
 * @returns {Object}
 */
export function parseTable(text) {
    // TODO: Implement table parsing
    console.log('[Stub] parseTable - not implemented');
    return { headers: [], rows: [], hasHeader: false };
}

/**
 * NumberingService for managing numbering.xml.
 * STUB: Not yet implemented.
 */
export class NumberingService {
    constructor() {
        console.log('[Stub] NumberingService - not implemented');
    }

    async initialize() {
        // TODO: Load numbering.xml
    }

    getOrCreateNumId(levelConfigs, existingContext = null) {
        // TODO: Implement numbering management
        return '1';  // Default numId
    }

    async commit() {
        // TODO: Write changes to numbering.xml
    }
}
