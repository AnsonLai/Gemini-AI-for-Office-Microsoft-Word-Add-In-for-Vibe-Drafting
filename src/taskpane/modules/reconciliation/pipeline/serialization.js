/**
 * OOXML Reconciliation Pipeline - Serialization
 * 
 * Converts patched run model back to OOXML with track changes.
 */

import { RunKind, escapeXml, getNextRevisionId } from '../core/types.js';
import { getApplicableFormatHints } from './markdown-processor.js';
import { serializeXml } from '../adapters/xml-adapter.js';
import { warn } from '../adapters/logger.js';
import { buildDocumentFragmentPackage } from '../services/package-builder.js';

/**
 * Serializes a patched run model to OOXML.
 * 
 * @param {import('../core/types.js').RunEntry[]} patchedModel - The patched run model
 * @param {Element|null} pPr - Paragraph properties element
 * @param {import('../core/types.js').FormatHint[]} [formatHints=[]] - Format hints
 * @param {import('../core/types.js').SerializationOptions} [options={}] - Serialization options
 * @returns {string} OOXML paragraph string (WITHOUT namespace - added by wrapper)
 */
export function serializeToOoxml(patchedModel, pPr, formatHints = [], options = {}) {
    const serializationOptions = normalizeSerializationOptions(options);
    const { author, generateRedlines } = serializationOptions;
    const paragraphs = [];
    let currentPPrXml = '';
    let currentRuns = [];

    // Helper to flush accumulated runs into a paragraph
    function flushParagraph() {
        if (currentRuns.length > 0 || paragraphs.length === 0) {
            // Build paragraph properties - handle both string and DOM element
            let pPrContent = '';
            if (currentPPrXml) {
                pPrContent = currentPPrXml.replace(/\s+xmlns:[^=]+="[^"]*"/g, '');
            } else if (pPr) {
                // Fallback to legacy pPr if no PARAGRAPH_START was seen
                if (typeof pPr === 'string') {
                    pPrContent = pPr;
                } else {
                    pPrContent = serializeXml(pPr);
                }
                pPrContent = pPrContent.replace(/\s+xmlns:[^=]+="[^"]*"/g, '');
            }
            paragraphs.push(`<w:p>${pPrContent}${currentRuns.join('')}</w:p>`);
            currentRuns = [];
        }
    }

    for (const item of patchedModel) {
        switch (item.kind) {
            case RunKind.PARAGRAPH_START:
                // Flush previous paragraph before starting a new one
                if (currentRuns.length > 0 || paragraphs.length > 0) {
                    flushParagraph();
                }
                currentPPrXml = item.pPrXml || '';
                break;

            case RunKind.TEXT:
                currentRuns.push(buildRunXmlWithHints(item, formatHints, serializationOptions));
                break;

            case RunKind.DELETION:
                if (generateRedlines) {
                    currentRuns.push(buildDeletionXml(item, serializationOptions));
                }
                // If redlines are disabled, we simply omit the deleted content
                break;

            case RunKind.INSERTION:
                if (generateRedlines) {
                    currentRuns.push(buildInsertionXml(item, formatHints, serializationOptions));
                } else {
                    // Treat insertion as a normal run when redlines are disabled
                    currentRuns.push(buildRunXmlWithHints(item, formatHints, serializationOptions));
                }
                break;

            case RunKind.BOOKMARK:
            case RunKind.HYPERLINK:
                // Pass through original XML - but strip any namespace declarations
                if (item.nodeXml) {
                    currentRuns.push(item.nodeXml.replace(/\s+xmlns:[^=]+="[^"]*"/g, ''));
                }
                break;

            case RunKind.CONTAINER_START:
                if (item.containerKind === 'sdt') {
                    currentRuns.push(`<w:sdt>${item.propertiesXml}<w:sdtContent>`);
                } else if (item.containerKind === 'smartTag') {
                    currentRuns.push(`<w:smartTag ${item.propertiesXml}>`);
                } else if (item.containerKind === 'hyperlink') {
                    const props = JSON.parse(item.propertiesXml);
                    const rIdAttr = props.rId ? ` r:id="${props.rId}"` : '';
                    const anchorAttr = props.anchor ? ` w:anchor="${props.anchor}"` : '';
                    currentRuns.push(`<w:hyperlink${rIdAttr}${anchorAttr}>`);
                }
                break;

            case RunKind.CONTAINER_END:
                if (item.containerKind === 'sdt') {
                    currentRuns.push(`</w:sdtContent></w:sdt>`);
                } else if (item.containerKind === 'smartTag') {
                    currentRuns.push(`</w:smartTag>`);
                } else if (item.containerKind === 'hyperlink') {
                    currentRuns.push(`</w:hyperlink>`);
                }
                break;

            default:
                warn('Unknown run kind:', item.kind);
        }
    }

    // Flush final paragraph
    flushParagraph();

    // Return all paragraphs WITHOUT namespace - wrapper will add it
    return paragraphs.join('');
}

/**
 * Normalizes serialization options to a single object contract.
 *
 * @param {import('../core/types.js').SerializationOptions|string|undefined|null} options - Raw options
 * @returns {import('../core/types.js').SerializationOptions}
 */
function normalizeSerializationOptions(options) {
    // Backward compatibility: accept legacy `font` string signature.
    if (typeof options === 'string') {
        return {
            author: 'Gemini AI',
            generateRedlines: true,
            font: options
        };
    }

    const normalized = options && typeof options === 'object' ? options : {};
    return {
        author: normalized.author ?? 'Gemini AI',
        generateRedlines: normalized.generateRedlines ?? true,
        font: normalized.font ?? null
    };
}

/**
 * Builds a run XML element, applying format hints if applicable.
 * 
 * @param {import('../core/types.js').RunEntry} item - Run entry
 * @param {import('../core/types.js').FormatHint[]} formatHints - Format hints
 * @param {import('../core/types.js').SerializationOptions} options - Serialization options
 * @returns {string}
 */
function buildRunXmlWithHints(item, formatHints, options = {}) {
    const applicableHints = getApplicableFormatHints(formatHints, item.startOffset, item.endOffset);
    const font = options?.font ?? null;
    let cleanRPr = item.rPrXml ? item.rPrXml.replace(/\s+xmlns:[^=]+="[^"]*"/g, '') : '';

    if (font) {
        cleanRPr = applyFont(cleanRPr, font);
    }

    if (applicableHints.length === 0) {
        // No formatting changes - use original rPr (strip namespace)
        return buildSimpleRun(item.text, cleanRPr);
    }

    // Split the run text at format boundaries and apply hints
    const runs = [];
    let pos = 0;
    const text = item.text;
    const baseOffset = item.startOffset;

    for (const hint of applicableHints) {
        const localStart = Math.max(0, hint.start - baseOffset);
        const localEnd = Math.min(text.length, hint.end - baseOffset);

        // Text before the hint
        if (localStart > pos) {
            runs.push(buildSimpleRun(text.slice(pos, localStart), cleanRPr));
        }

        // Formatted text
        const formattedRPr = injectFormatting(cleanRPr, hint.format);
        runs.push(buildSimpleRun(text.slice(localStart, localEnd), formattedRPr));
        pos = localEnd;
    }

    // Remaining text after last hint
    if (pos < text.length) {
        runs.push(buildSimpleRun(text.slice(pos), cleanRPr));
    }

    return runs.join('');
}

/**
 * Builds a simple w:r element.
 * 
 * @param {string} text - Text content
 * @param {string} rPrXml - Run properties XML
 * @returns {string}
 */
function buildSimpleRun(text, rPrXml) {
    if (!text) return '';
    const rPr = rPrXml || '';
    return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

/**
 * Builds a deletion (w:del) element.
 * 
 * @param {import('../core/types.js').RunEntry} item - Deletion entry
 * @param {import('../core/types.js').SerializationOptions} options - Serialization options
 * @returns {string}
 */
function buildDeletionXml(item, options = {}) {
    const author = options.author ?? 'Gemini AI';
    const font = options.font ?? null;
    const revId = getNextRevisionId();
    const date = new Date().toISOString();
    let rPr = item.rPrXml ? item.rPrXml.replace(/\s+xmlns:[^=]+="[^"]*"/g, '') : '';

    if (font) {
        rPr = applyFont(rPr, font);
    }

    return `<w:del w:id="${revId}" w:author="${escapeXml(author)}" w:date="${date}">` +
        `<w:r>${rPr}<w:delText xml:space="preserve">${escapeXml(item.text)}</w:delText></w:r>` +
        `</w:del>`;
}

/**
 * Builds an insertion (w:ins) element.
 * 
 * @param {import('../core/types.js').RunEntry} item - Insertion entry
 * @param {import('../core/types.js').FormatHint[]} formatHints - Format hints
 * @param {import('../core/types.js').SerializationOptions} options - Serialization options
 * @returns {string}
 */
function buildInsertionXml(item, formatHints, options = {}) {
    const author = options.author ?? 'Gemini AI';
    const font = options.font ?? null;
    const revId = getNextRevisionId();
    const date = new Date().toISOString();

    // Build the inner run content with format hints
    const applicableHints = getApplicableFormatHints(formatHints, item.startOffset, item.endOffset);
    let innerContent = '';
    let cleanRPr = item.rPrXml ? item.rPrXml.replace(/\s+xmlns:[^=]+="[^"]*"/g, '') : '';

    if (font) {
        cleanRPr = applyFont(cleanRPr, font);
    }

    if (applicableHints.length === 0) {
        innerContent = buildSimpleRun(item.text, cleanRPr);
    } else {
        // Apply format hints
        let pos = 0;
        const text = item.text;
        const baseOffset = item.startOffset;

        for (const hint of applicableHints) {
            const localStart = Math.max(0, hint.start - baseOffset);
            const localEnd = Math.min(text.length, hint.end - baseOffset);

            if (localStart > pos) {
                innerContent += buildSimpleRun(text.slice(pos, localStart), cleanRPr);
            }

            const formattedRPr = injectFormatting(cleanRPr, hint.format);
            innerContent += buildSimpleRun(text.slice(localStart, localEnd), formattedRPr);
            pos = localEnd;
        }

        if (pos < text.length) {
            innerContent += buildSimpleRun(text.slice(pos), cleanRPr);
        }
    }

    return `<w:ins w:id="${revId}" w:author="${escapeXml(author)}" w:date="${date}">` +
        innerContent +
        `</w:ins>`;
}

/**
 * Applies a font to run properties XML.
 * 
 * @param {string} baseRPrXml - Base run properties
 * @param {string} font - Font name
 * @returns {string}
 */
function applyFont(baseRPrXml, font) {
    if (!font) return baseRPrXml;

    // Extract existing content from rPr
    let content = '';
    if (baseRPrXml) {
        content = baseRPrXml.replace(/<\/?w:rPr[^>]*>/g, '');
    }

    // Replace or add rFonts
    if (content.includes('<w:rFonts')) {
        content = content.replace(/<w:rFonts[^>]*\/>/, `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>`);
    } else {
        content = `<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>` + content;
    }

    return `<w:rPr>${content}</w:rPr>`;
}

/**
 * Injects formatting into run properties XML.
 * 
 * @param {string} baseRPrXml - Base run properties
 * @param {Object} format - Format flags (bold, italic, underline, strikethrough)
 * @returns {string}
 */
function injectFormatting(baseRPrXml, format) {
    if (!format || Object.keys(format).length === 0) {
        return baseRPrXml;
    }

    // Extract existing content from rPr
    let content = '';
    if (baseRPrXml) {
        content = baseRPrXml.replace(/<\/?w:rPr[^>]*>/g, '');
    }

    // Add new formatting elements
    if (format.bold && !content.includes('<w:b')) {
        content = '<w:b/>' + content;
    }
    if (format.italic && !content.includes('<w:i')) {
        content = '<w:i/>' + content;
    }
    if (format.underline && !content.includes('<w:u')) {
        content = '<w:u w:val="single"/>' + content;
    }
    if (format.strikethrough && !content.includes('<w:strike')) {
        content = '<w:strike/>' + content;
    }

    return `<w:rPr>${content}</w:rPr>`;
}

/**
 * Wraps OOXML paragraph content for Word's insertOoxml API.
 * Must include both the document part AND the relationships part.
 * 
 * @param {string} paragraphXml - The paragraph XML (without namespace declarations)
 * @param {import('../core/types.js').DocumentFragmentOptions|boolean} [options={}] - Fragment options
 * @returns {string} Complete OOXML package for insertOoxml
 */
export function wrapInDocumentFragment(paragraphXml, options = {}) {
    const normalizedOptions = normalizeFragmentOptions(options);
    return buildDocumentFragmentPackage(paragraphXml, normalizedOptions);
}

/**
 * Normalizes wrapper options to object form.
 *
 * @param {import('../core/types.js').DocumentFragmentOptions|boolean|undefined|null} options - Raw options
 * @returns {import('../core/types.js').DocumentFragmentOptions}
 */
function normalizeFragmentOptions(options) {
    if (typeof options === 'boolean') {
        return { includeNumbering: options };
    }

    if (!options || typeof options !== 'object') {
        return {};
    }

    return {
        includeNumbering: options.includeNumbering ?? false,
        numberingXml: options.numberingXml ?? null,
        appendTrailingParagraph: options.appendTrailingParagraph ?? true
    };
}

