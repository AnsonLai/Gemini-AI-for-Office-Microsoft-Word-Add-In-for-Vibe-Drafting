/**
 * Shared paragraph-targeting helpers for standalone/add-in consumers.
 *
 * This module centralizes target parsing and matching used by callers that
 * apply per-paragraph operations (for example chat redlines/comments/highlights).
 */

function toArray(nodeList) {
    return Array.from(nodeList || []);
}

function getElementsByLocalName(node, localName) {
    if (!node) return [];

    if (typeof node.getElementsByTagNameNS === 'function') {
        const namespaced = toArray(node.getElementsByTagNameNS('*', localName));
        if (namespaced.length > 0) return namespaced;
    }

    if (typeof node.getElementsByTagName !== 'function') return [];

    const prefixed = toArray(node.getElementsByTagName(`w:${localName}`));
    if (prefixed.length > 0) return prefixed;

    return toArray(node.getElementsByTagName(localName));
}

function toParagraphText(paragraph) {
    const textNodes = getElementsByLocalName(paragraph, 't');
    return textNodes.map(node => node.textContent || '').join('');
}

/**
 * Reads visible text from a paragraph by concatenating `w:t` nodes.
 *
 * @param {Element|null|undefined} paragraph - OOXML paragraph node
 * @returns {string}
 */
export function getParagraphText(paragraph) {
    if (!paragraph) return '';
    return toParagraphText(paragraph);
}

/**
 * Returns body paragraphs for a document, or all paragraphs as fallback.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document root
 * @returns {Element[]}
 */
export function getDocumentParagraphNodes(xmlDoc) {
    if (!xmlDoc) return [];
    const bodies = getElementsByLocalName(xmlDoc, 'body');
    const searchRoot = bodies.length > 0 ? bodies[0] : xmlDoc;
    return getElementsByLocalName(searchRoot, 'p');
}

/**
 * Normalizes whitespace for paragraph-comparison matching.
 *
 * @param {string} text - Input text
 * @returns {string}
 */
export function normalizeWhitespaceForTargeting(text) {
    return String(text || '').replace(/\s+/g, ' ').trim();
}

/**
 * Parses paragraph references such as `P12`, `[P12]`, `12`, or `P12.3`.
 *
 * @param {string|number|null|undefined} rawValue - Reference input
 * @returns {number|null}
 */
export function parseParagraphReference(rawValue) {
    if (rawValue == null) return null;
    if (typeof rawValue === 'number' && Number.isInteger(rawValue) && rawValue > 0) return rawValue;

    const text = String(rawValue).trim();
    if (!text) return null;

    const prefixed = text.match(/^\[?P(\d+)(?:\.\d+)?\]?$/i);
    if (prefixed) return Number.parseInt(prefixed[1], 10);

    const numeric = text.match(/^(\d+)$/);
    if (numeric) return Number.parseInt(numeric[1], 10);

    return null;
}

/**
 * Removes leading paragraph labels (for example `[P12]`) from text fields.
 *
 * @param {string|null|undefined} text - Input text
 * @returns {string}
 */
export function stripLeadingParagraphMarker(text) {
    if (text == null) return '';
    return String(text).replace(/^\s*\[P\d+(?:\.\d+)?\]\s*/i, '').trim();
}

/**
 * Splits a leading paragraph label from text.
 *
 * @param {string|null|undefined} text - Input text
 * @returns {{ text: string, targetRef: number|null }}
 */
export function splitLeadingParagraphMarker(text) {
    const raw = String(text || '');
    const marker = raw.match(/^\s*\[P(\d+)(?:\.\d+)?\]\s*/i);
    if (!marker) return { text: raw.trim(), targetRef: null };

    return {
        text: raw.replace(/^\s*\[P\d+(?:\.\d+)?\]\s*/i, '').trim(),
        targetRef: Number.parseInt(marker[1], 10)
    };
}

/**
 * Resolves a paragraph by 1-based paragraph index.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {number|null|undefined} targetRef - 1-based paragraph number
 * @returns {Element|null}
 */
export function findParagraphByReference(xmlDoc, targetRef) {
    if (!Number.isInteger(targetRef) || targetRef < 1) return null;
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    return paragraphs[targetRef - 1] || null;
}

/**
 * Finds paragraph by exact/normalized text equality.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {string} targetText - Target paragraph text
 * @returns {Element|null}
 */
export function findParagraphByStrictText(xmlDoc, targetText) {
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    const normalizedTarget = String(targetText || '').trim();
    if (!normalizedTarget) return null;

    const exact = paragraphs.find(p => getParagraphText(p).trim() === normalizedTarget);
    if (exact) return exact;

    const normTarget = normalizeWhitespaceForTargeting(normalizedTarget);
    return paragraphs.find(p => normalizeWhitespaceForTargeting(getParagraphText(p)) === normTarget) || null;
}

/**
 * Finds paragraph by strict match, then fuzzy fallback heuristics.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {string} targetText - Target paragraph text
 * @param {{ onInfo?: (msg:string)=>void }} [options] - Optional logger callbacks
 * @returns {Element|null}
 */
export function findParagraphByBestTextMatch(xmlDoc, targetText, options = {}) {
    const onInfo = typeof options.onInfo === 'function' ? options.onInfo : () => {};
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    const normalizedTarget = String(targetText || '').trim();
    if (!normalizedTarget) return null;

    const strictMatch = findParagraphByStrictText(xmlDoc, normalizedTarget);
    if (strictMatch) return strictMatch;

    const normTarget = normalizeWhitespaceForTargeting(normalizedTarget);

    const startsWithMatch = paragraphs.find(p => {
        const paragraphText = normalizeWhitespaceForTargeting(getParagraphText(p));
        return paragraphText.length > 10 && normTarget.startsWith(paragraphText);
    });
    if (startsWithMatch) {
        onInfo(`[Fuzzy] Prefix match (target starts with paragraph): "${getParagraphText(startsWithMatch).trim().slice(0, 60)}..."`);
        return startsWithMatch;
    }

    const containsMatch = paragraphs.find(p => {
        const paragraphText = normalizeWhitespaceForTargeting(getParagraphText(p));
        return paragraphText.length > 15 && normTarget.includes(paragraphText);
    });
    if (containsMatch) {
        onInfo(`[Fuzzy] Contains match: "${getParagraphText(containsMatch).trim().slice(0, 60)}..."`);
        return containsMatch;
    }

    let bestScore = 0;
    let bestParagraph = null;
    const targetWords = new Set(normTarget.toLowerCase().split(/\s+/).filter(word => word.length > 2));
    for (const paragraph of paragraphs) {
        const paragraphText = getParagraphText(paragraph).trim();
        if (!paragraphText) continue;

        const paragraphWords = normalizeWhitespaceForTargeting(paragraphText)
            .toLowerCase()
            .split(/\s+/)
            .filter(word => word.length > 2);
        const overlap = paragraphWords.filter(word => targetWords.has(word)).length;
        const score = overlap / Math.max(targetWords.size, 1);
        if (score > bestScore && score > 0.5) {
            bestScore = score;
            bestParagraph = paragraph;
        }
    }

    if (bestParagraph) {
        onInfo(`[Fuzzy] Best word-overlap match (${(bestScore * 100).toFixed(0)}%): "${getParagraphText(bestParagraph).trim().slice(0, 60)}..."`);
        return bestParagraph;
    }

    return null;
}

/**
 * Resolves a target paragraph from `targetRef` + `targetText`.
 *
 * Resolution order:
 * 1) `targetRef` when provided and valid
 * 2) strict text match
 * 3) fuzzy text match
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {{
 *   targetText?: string,
 *   targetRef?: string|number|null,
 *   opType?: string,
 *   onInfo?: (msg:string)=>void,
 *   onWarn?: (msg:string)=>void
 * }} options - Resolution options
 * @returns {{ paragraph: Element, resolvedBy: 'ref'|'strict_text'|'fuzzy_text' }}
 */
export function resolveTargetParagraph(xmlDoc, options = {}) {
    const onInfo = typeof options.onInfo === 'function' ? options.onInfo : () => {};
    const onWarn = typeof options.onWarn === 'function' ? options.onWarn : () => {};
    const opType = options.opType || 'operation';
    const cleanTargetText = String(options.targetText || '').trim();
    const parsedRef = parseParagraphReference(options.targetRef);

    if (parsedRef) {
        const byRef = findParagraphByReference(xmlDoc, parsedRef);
        if (byRef) {
            if (cleanTargetText) {
                const strictMatch = findParagraphByStrictText(xmlDoc, cleanTargetText);
                if (strictMatch && strictMatch !== byRef) {
                    onInfo(`[Target] [P${parsedRef}] disambiguated duplicate target text for ${opType}.`);
                }

                const byRefText = getParagraphText(byRef).trim();
                if (normalizeWhitespaceForTargeting(byRefText) !== normalizeWhitespaceForTargeting(cleanTargetText)) {
                    onInfo(`[Target] Using [P${parsedRef}] fallback for ${opType}; target text drifted.`);
                }
            } else {
                onInfo(`[Target] Using [P${parsedRef}] fallback for ${opType}.`);
            }
            return { paragraph: byRef, resolvedBy: 'ref' };
        }

        onWarn(`[WARN] Target reference [P${parsedRef}] not found; falling back to text matching for ${opType}.`);
    }

    if (cleanTargetText) {
        const strictMatch = findParagraphByStrictText(xmlDoc, cleanTargetText);
        if (strictMatch) return { paragraph: strictMatch, resolvedBy: 'strict_text' };

        const fuzzyMatch = findParagraphByBestTextMatch(xmlDoc, cleanTargetText, { onInfo });
        if (fuzzyMatch) return { paragraph: fuzzyMatch, resolvedBy: 'fuzzy_text' };
    }

    if (cleanTargetText) throw new Error(`Target paragraph not found: "${cleanTargetText}"`);
    if (parsedRef) throw new Error(`Target paragraph reference not found: [P${parsedRef}]`);
    throw new Error('Operation target missing: provide "target" text or "targetRef" ([P#]).');
}
