import { advanceOffsetForParagraphBoundary } from '../core/paragraph-offset-policy.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function isWordElement(node, localName) {
    if (!node || node.nodeType !== 1) return false;
    if (node.namespaceURI === NS_W && node.localName === localName) return true;
    const nodeName = String(node.nodeName || '');
    return nodeName === `w:${localName}` || nodeName === localName;
}

/**
 * Paragraph targeting helpers for format-only operations.
 *
 * Encapsulates paragraph text reconstruction and matching logic used to map
 * AI-provided text ranges to the correct paragraph/spans in OOXML.
 */

/**
 * Builds paragraph metadata (text, spans, offsets) used for format-only changes.
 *
 * @param {Document} xmlDoc - XML document (unused, kept for signature compatibility)
 * @param {Element[]} paragraphs - Paragraph elements
 * @param {Array} textSpans - Extracted text spans
 * @returns {Array}
 */
export function buildParagraphInfos(xmlDoc, paragraphs, textSpans) {
    void xmlDoc;
    const spansByParagraph = new Map();
    for (const span of textSpans) {
        if (!span || !span.paragraph) continue;
        if (!spansByParagraph.has(span.paragraph)) {
            spansByParagraph.set(span.paragraph, []);
        }
        spansByParagraph.get(span.paragraph).push(span);
    }

    const infos = [];
    let runningOffset = 0;

    paragraphs.forEach((p, index) => {
        const spans = (spansByParagraph.get(p) || []).slice().sort((a, b) => a.charStart - b.charStart);
        const text = buildParagraphTextFromSpans(spans);
        const normalizedText = normalizeParagraphComparisonText(text);
        const normalizedTrim = normalizedText.trim();

        infos.push({
            paragraph: p,
            spans,
            text,
            normalizedText,
            normalizedTrim,
            startOffset: runningOffset
        });

        runningOffset += normalizedText.length;
        runningOffset = advanceOffsetForParagraphBoundary(runningOffset, index, paragraphs.length);
    });

    return infos;
}

/**
 * Finds a paragraph info entry that matches the provided text.
 *
 * @param {Array} paragraphInfos - Paragraph metadata
 * @param {string} originalText - Original text
 * @returns {Object|null}
 */
export function findMatchingParagraphInfo(paragraphInfos, originalText) {
    if (!originalText) return null;

    const normalizedOriginal = normalizeParagraphComparisonText(originalText);
    const normalizedTrim = normalizedOriginal.trim();
    if (!normalizedTrim) return null;

    for (const info of paragraphInfos) {
        if (info.normalizedText === normalizedOriginal) {
            return info;
        }
    }

    for (const info of paragraphInfos) {
        if (info.normalizedTrim === normalizedTrim) {
            return info;
        }
    }

    return null;
}

/**
 * Finds the best target paragraph and offset for format-only operations.
 *
 * @param {Array} paragraphInfos - Paragraph metadata
 * @param {string} originalText - Original text used for matching
 * @returns {{ targetInfo: Object|null, matchOffset: number }}
 */
export function findTargetParagraphInfo(paragraphInfos, originalText) {
    const normalizedOriginalFull = normalizeParagraphComparisonText(originalText);
    const normalizedOriginalTrim = normalizedOriginalFull.trim();

    let targetInfo = null;
    let matchOffset = 0;

    for (const info of paragraphInfos) {
        if (info.normalizedText === normalizedOriginalFull) {
            targetInfo = info;
            return { targetInfo, matchOffset };
        }
    }

    if (normalizedOriginalTrim.length > 0) {
        for (const info of paragraphInfos) {
            if (info.normalizedTrim === normalizedOriginalTrim) {
                targetInfo = info;
                return { targetInfo, matchOffset };
            }
        }
    }

    if (normalizedOriginalTrim.length > 0) {
        const docPlain = paragraphInfos.map(info => info.normalizedText).join('\n');
        const idx = docPlain.indexOf(normalizedOriginalTrim);
        if (idx !== -1) {
            for (const info of paragraphInfos) {
                const start = info.startOffset;
                const length = info.normalizedText.length;
                if (idx >= start && idx <= start + length) {
                    targetInfo = info;
                    matchOffset = idx - start;
                    break;
                }
            }
        }
    }

    // Fallback: single-paragraph scope where extracted span text is only a subset
    // of original paragraph text (common when prior tracked wrappers contain text
    // that is not represented in direct run-span extraction).
    if (!targetInfo && paragraphInfos.length === 1 && normalizedOriginalFull.length > 0) {
        const onlyInfo = paragraphInfos[0];
        const paragraphTrim = onlyInfo.normalizedTrim || '';
        if (paragraphTrim.length > 0) {
            const subsetIndex = normalizedOriginalFull.indexOf(paragraphTrim);
            if (subsetIndex >= 0) {
                targetInfo = onlyInfo;
                matchOffset = -subsetIndex;
            }
        }
    }

    return { targetInfo, matchOffset };
}

/**
 * Walks up the DOM to find the containing paragraph for a run.
 *
 * @param {Node} node - Starting node
 * @returns {Element|null}
 */
export function getContainingParagraph(node) {
    let current = node;
    while (current) {
        if (isWordElement(current, 'p')) return current;
        current = current.parentNode;
    }
    return null;
}

/**
 * Reconstructs human-readable text from spans to align with Word paragraph text.
 *
 * @param {Array} spans - Span collection
 * @returns {string}
 */
function buildParagraphTextFromSpans(spans) {
    if (!spans || spans.length === 0) return '';

    let text = '';
    for (const span of spans) {
        if (!span || !span.textElement) continue;
        const textElement = span.textElement;
        if (isWordElement(textElement, 't')) {
            text += span.textElement.textContent || '';
        } else if (isWordElement(textElement, 'tab')) {
            text += '\t';
        } else if (isWordElement(textElement, 'br') || isWordElement(textElement, 'cr')) {
            text += '\n';
        } else if (isWordElement(textElement, 'noBreakHyphen')) {
            text += '\u2011';
        }
    }
    return text;
}

/**
 * Normalizes paragraph text for comparisons (handles carriage returns and NBSP).
 *
 * @param {string} text - Text input
 * @returns {string}
 */
function normalizeParagraphComparisonText(text) {
    if (!text) return '';
    return text
        .replace(/\r/g, '\n')
        .replace(/\u00a0/g, ' ');
}
