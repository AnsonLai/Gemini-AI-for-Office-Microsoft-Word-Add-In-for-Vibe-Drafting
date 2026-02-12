/**
 * OOXML Reconciliation Pipeline - Main Pipeline
 * 
 * Orchestrates the reconciliation process from OOXML input to output.
 */

import { ingestOoxml } from './ingestion.js';
import { preprocessMarkdown } from './markdown-processor.js';
import { isListTargetLoose, isListTargetStrict } from './list-markers.js';
import { computeWordLevelDiffOps } from './diff-engine.js';
import { splitRunsAtDiffBoundaries, applyPatches } from './patching.js';
import { serializeToOoxml, wrapInDocumentFragment } from './serialization.js';
import { RunKind } from '../core/types.js';
import { NumberingService } from '../services/numbering-service.js';
import { detectNumberingContext } from './ingestion.js';
import { generateTableOoxml } from '../services/table-reconciliation.js';
import { executeListGeneration, detectIndentationStep } from './list-generation.js';
import { detectContentType, parseListItems, parseTable } from './content-analysis.js';
import { createParser } from '../adapters/xml-adapter.js';
import { log, error as logError } from '../adapters/logger.js';
import { getFirstElementByTagNS, getXmlParseError } from '../core/xml-query.js';

const WEB_PLATFORM_NAMES = new Set(['officeonline', 'officeweb', 'web']);

function detectOfficePlatform() {
    if (typeof Office === 'undefined' || !Office?.context?.platform) {
        return null;
    }
    return String(Office.context.platform);
}

function isWebPlatform(platform) {
    if (!platform) return false;
    return WEB_PLATFORM_NAMES.has(String(platform).toLowerCase());
}

function isProductionBuild() {
    return typeof process !== 'undefined' && process?.env?.NODE_ENV === 'production';
}

function yieldToEventLoop() {
    return new Promise(resolve => setTimeout(resolve, 0));
}

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
        this.validationMode = options.validationMode ?? 'auto';
        this.numberingService = options.numberingService || new NumberingService();
        this.font = options.font || null;
        this.platform = options.platform ?? detectOfficePlatform();
        this.isWebPlatform = options.isWebPlatform ?? isWebPlatform(this.platform);
        this.enableEventLoopYielding = options.enableEventLoopYielding ?? this.isWebPlatform;
        this.yieldRunThreshold = options.yieldRunThreshold ?? 50;
        this.yieldCharThreshold = options.yieldCharThreshold ?? 5000;
        this.disableSemanticCleanupOverChars = options.disableSemanticCleanupOverChars ?? (this.isWebPlatform ? 8000 : Number.POSITIVE_INFINITY);
    }

    /**
     * Executes the reconciliation pipeline.
     * 
     * @param {string} originalOoxml - Original OOXML paragraph content
     * @param {string} newText - New text with optional markdown formatting
     * @param {{ xmlDoc?: Document|null }} [options={}] - Optional execution options
     * @returns {Promise<import('../core/types.js').ReconciliationResult>}
     */
    async execute(originalOoxml, newText, options = {}) {
        const warnings = [];

        try {
            // Stage 1: Ingest OOXML
            const doc = options.xmlDoc || (() => {
                const parser = createParser();
                return parser.parseFromString(originalOoxml, 'application/xml');
            })();
            const pElement = getFirstElementByTagNS(doc, '*', 'p');

            const { runModel, acceptedText, pPr } = ingestOoxml(originalOoxml, { xmlDoc: doc });
            const numberingContext = pElement ? detectNumberingContext(pElement) : null;

            log(`[Reconcile] Ingested ${runModel.length} runs, ${acceptedText.length} chars, numbering:`, numberingContext);
            await this.maybeYield(runModel.length, Math.max(acceptedText.length, newText?.length || 0));

            // Stage 2: Preprocess markdown
            const { cleanText, formatHints } = preprocessMarkdown(newText);
            log(`[Reconcile] Preprocessed: ${formatHints.length} format hints`);
            await this.maybeYield(runModel.length, Math.max(acceptedText.length, cleanText.length));

            // Detect list-target content before any no-op short-circuit.
            // Structural conversion may still be required even when text is identical
            // (for example plain "A./B./C." lines -> true Word numbered list paragraphs).
            const isTargetListStrict = isListTargetStrict(cleanText);
            const isTargetListLoose = isListTargetLoose(cleanText);
            const isTargetList = isTargetListStrict || isTargetListLoose;
            if (!isTargetListStrict && isTargetListLoose) {
                log('[Reconcile] List-target detected via loose marker parsing; bypassing no-op short-circuit for structural conversion.');
            }

            // Early exit if no change
            if (acceptedText === cleanText && formatHints.length === 0 && !isTargetList) {
                log('[Reconcile] No changes detected');
                return {
                    ooxml: originalOoxml,
                    isValid: true,
                    warnings: ['No changes detected']
                };
            }

            // Stage 3: Compute word-level diff
            const shouldCleanupSemantic = Math.max(acceptedText.length, cleanText.length) < this.disableSemanticCleanupOverChars;
            const diffOps = computeWordLevelDiffOps(acceptedText, cleanText, {
                cleanupSemantic: shouldCleanupSemantic
            });
            if (!shouldCleanupSemantic) {
                log('[Reconcile] Skipping semantic diff cleanup for large web payload');
            }
            await this.maybeYield(runModel.length, Math.max(acceptedText.length, cleanText.length));

            // Count actual paragraph elements ingested
            const paragraphCount = runModel.filter(r => r.kind === RunKind.PARAGRAPH_START).length;

            log(`[Reconcile] isTargetList: ${isTargetList}, paragraphCount: ${paragraphCount}`);

            // If target is a list, always use list generation logic
            // This handles both expansion (1 para -> N items) and conversion (N paras -> M items)
            if (isTargetList) {
                log('[Reconcile] 🎯 ENTERING LIST GENERATION PATH');
                log(`[Reconcile] cleanText preview: ${cleanText.substring(0, 100)}...`);
                log(`[Reconcile] acceptedText preview: ${acceptedText.substring(0, 100)}...`);
                return this.executeListGeneration(cleanText, numberingContext, runModel);
            }

            log(`[Reconcile] Computed ${diffOps.length} diff operations`);

            // Stage 4: Pre-split runs at boundaries
            const splitModel = splitRunsAtDiffBoundaries(runModel, diffOps);
            log(`[Reconcile] Split into ${splitModel.length} runs`);

            // Stage 5: Apply patches
            const patchedModel = applyPatches(splitModel, diffOps, {
                generateRedlines: this.generateRedlines,
                author: this.author,
                formatHints,
                numberingService: this.numberingService
            });
            log(`[Reconcile] Patched model has ${patchedModel.length} runs`);
            await this.maybeYield(patchedModel.length, Math.max(acceptedText.length, cleanText.length));

            // Stage 6: Serialize to OOXML
            const resultOoxml = serializeToOoxml(patchedModel, pPr, formatHints, {
                author: this.author,
                generateRedlines: this.generateRedlines
            });

            // Stage 7: Basic validation
            if (this.shouldRunValidation()) {
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
            logError('[Reconcile] Pipeline error:', error);
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
            const parser = createParser();
            const doc = parser.parseFromString(wrappedXml, 'application/xml');

            const parseError = getXmlParseError(doc);
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
     * Decides if basic output validation should run for this pipeline instance.
     *
     * Modes:
     * - `always`: validate whenever `validateOutput` is true
     * - `never`: never validate
     * - `auto` (default): skip only in production web runtime
     *
     * @returns {boolean}
     */
    shouldRunValidation() {
        if (!this.validateOutput) return false;

        if (this.validationMode === 'always') return true;
        if (this.validationMode === 'never') return false;

        // Auto mode: avoid extra parse round-trips in Word Online production.
        return !(this.isWebPlatform && isProductionBuild());
    }

    /**
     * Yields to the event loop for large operations to keep web UI responsive.
     *
     * @param {number} runCount - Run model size
     * @param {number} charCount - Text size
     * @returns {Promise<void>}
     */
    async maybeYield(runCount, charCount) {
        if (!this.enableEventLoopYielding) return;
        if (runCount <= this.yieldRunThreshold && charCount <= this.yieldCharThreshold) return;
        await yieldToEventLoop();
    }

    /**
     * Wraps the reconciled content for document insertion.
     * 
     * @param {string} ooxml - Reconciled OOXML paragraph
     * @param {import('../core/types.js').DocumentFragmentOptions|boolean} [options={}] - Fragment options
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
        return executeListGeneration({
            cleanText,
            numberingContext,
            originalRunModel,
            originalText,
            generateRedlines: this.generateRedlines,
            author: this.author,
            font: this.font,
            numberingService: this.numberingService
        });
    }

    /**
     * Heuristically detects the indentation step (number of spaces or tabs per level).
     * 
     * @param {string[]} lines - Array of lines
     * @returns {number} The detected step (defaulting to 2)
     */
    detectIndentationStep(lines) {
        return detectIndentationStep(lines);
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
export { detectContentType, parseListItems, parseTable };

