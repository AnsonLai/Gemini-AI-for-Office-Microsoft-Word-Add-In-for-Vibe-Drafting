/**
 * Reconstruction mapper.
 *
 * Builds paragraph/property/sentinel mappings and indexed lookups used by reconstruction writing.
 */

import { appendParagraphBoundary } from '../core/paragraph-offset-policy.js';
import { getDocumentParagraphs } from './format-extraction.js';
import { getElementsByTag, getFirstElementByTag } from '../core/xml-query.js';

function createRangeCursorLookup(ranges) {
    let cursor = 0;
    return {
        at(index) {
            while (cursor < ranges.length && ranges[cursor].end <= index) {
                cursor++;
            }
            const match = ranges[cursor];
            if (!match) return null;
            if (match.start <= index && index < match.end) return match;
            return null;
        }
    };
}

function indexSentinelsByStart(sentinelMap) {
    const sentinelMapByStart = new Map();
    sentinelMap.forEach(sentinel => {
        if (!sentinelMapByStart.has(sentinel.start)) {
            sentinelMapByStart.set(sentinel.start, []);
        }
        sentinelMapByStart.get(sentinel.start).push(sentinel);
    });
    return sentinelMapByStart;
}

/**
 * Builds reconstruction mapping and cursor-based lookup helpers.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} modifiedText - Modified text
 * @returns {{
 *   paragraphs: Element[],
 *   body: Element|Document,
 *   paragraphMap: Array<{start:number,end:number,pPr:Element|null,container:Node}>,
 *   paragraphStarts: Set<number>,
 *   propertyMap: Array<{start:number,end:number,rPr:Element|null,wrapper?:Element}>,
 *   sentinelMap: Array<Object>,
 *   sentinelMapByStart: Map<number, Object[]>,
 *   referenceMap: Map<string, Node>,
 *   tokenToCharMap: Map<string, string>,
 *   containerFragments: Map<Node, DocumentFragment>,
 *   replacementContainers: Map<Node, Node>,
 *   originalFullText: string,
 *   processedModifiedText: string,
 *   getParagraphInfo: (index:number) => {start:number,end:number,pPr:Element|null,container:Node},
 *   getRunProperties: (index:number) => {rPr:Element|null,wrapper?:Element},
 *   getPropertySpanLength: (index:number,maxLength:number) => number,
 *   isParagraphStart: (index:number) => boolean
 * }}
 */
export function buildReconstructionMapping(xmlDoc, modifiedText) {
    const rootElement = xmlDoc.documentElement;
    const isBodyRoot = rootElement.nodeName === 'w:body' || rootElement.nodeName.endsWith(':package');
    const paragraphs = getDocumentParagraphs(xmlDoc);

    let body = getFirstElementByTag(xmlDoc, 'w:body');
    if (!body && isBodyRoot) body = rootElement;

    let originalFullText = '';
    const propertyMap = [];
    const paragraphMap = [];
    const sentinelMap = [];
    const referenceMap = new Map();
    const tokenToCharMap = new Map();
    let nextCharCode = 0xe000;
    const uniqueContainers = new Set();

    paragraphs.forEach((paragraph, paragraphIndex) => {
        const paragraphStart = originalFullText.length;

        Array.from(paragraph.childNodes).forEach(child => {
            originalFullText = processChildNode(
                child,
                originalFullText,
                propertyMap,
                sentinelMap,
                referenceMap,
                tokenToCharMap,
                nextCharCode
            );
            if (referenceMap.size > tokenToCharMap.size) {
                nextCharCode++;
            }
        });

        originalFullText = appendParagraphBoundary(originalFullText, paragraphIndex, paragraphs.length);

        const paragraphEnd = originalFullText.length;
        const pPr = getFirstElementByTag(paragraph, 'w:pPr');
        const container = paragraph.parentNode;
        if (container) uniqueContainers.add(container);

        paragraphMap.push({
            start: paragraphStart,
            end: paragraphEnd,
            pPr,
            container: container || body
        });
    });

    let processedModifiedText = modifiedText;
    tokenToCharMap.forEach((char, tokenString) => {
        const escapedToken = tokenString.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
        processedModifiedText = processedModifiedText.replace(new RegExp(escapedToken, 'g'), char);
    });

    const containerFragments = new Map();
    uniqueContainers.forEach(container => {
        containerFragments.set(container, xmlDoc.createDocumentFragment());
    });
    if (body && !containerFragments.has(body)) {
        containerFragments.set(body, xmlDoc.createDocumentFragment());
    }
    if (!containerFragments.has(xmlDoc)) {
        containerFragments.set(xmlDoc, xmlDoc.createDocumentFragment());
    }

    const replacementContainers = new Map();
    const paragraphStarts = new Set(paragraphMap.map(paragraph => paragraph.start));
    const paragraphLookup = createRangeCursorLookup(paragraphMap);
    const propertyLookup = createRangeCursorLookup(propertyMap);
    const sentinelMapByStart = indexSentinelsByStart(sentinelMap);

    const getParagraphInfo = (index) => {
        const match = paragraphLookup.at(index);
        if (match) return match;
        if (paragraphMap.length > 0) return paragraphMap[paragraphMap.length - 1];
        return { start: 0, end: 0, pPr: null, container: body || xmlDoc };
    };

    const getRunProperties = (index) => {
        const match = propertyLookup.at(index);
        return match ? { rPr: match.rPr, wrapper: match.wrapper } : { rPr: null };
    };

    const getPropertySpanLength = (index, maxLength) => {
        const match = propertyLookup.at(index);
        if (!match) return 1;
        return Math.min(match.end - index, maxLength);
    };

    return {
        paragraphs,
        body: body || xmlDoc,
        paragraphMap,
        paragraphStarts,
        propertyMap,
        sentinelMap,
        sentinelMapByStart,
        referenceMap,
        tokenToCharMap,
        containerFragments,
        replacementContainers,
        originalFullText,
        processedModifiedText,
        getParagraphInfo,
        getRunProperties,
        getPropertySpanLength,
        isParagraphStart: index => paragraphStarts.has(index)
    };
}

function processChildNode(child, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, nextCharCode) {
    if (child.nodeName === 'w:r') {
        return processRunForReconstruction(child, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, nextCharCode);
    }
    if (child.nodeName === 'w:hyperlink') {
        return processHyperlinkForReconstruction(child, originalFullText, propertyMap);
    }
    if (['w:sdt', 'w:oMath', 'm:oMath', 'w:bookmarkStart', 'w:bookmarkEnd'].includes(child.nodeName)) {
        sentinelMap.push({ start: originalFullText.length, node: child });
        return originalFullText + '\uFFFC';
    }
    if (['w:commentRangeStart', 'w:commentRangeEnd'].includes(child.nodeName)) {
        sentinelMap.push({ start: originalFullText.length, node: child, isCommentMarker: true });
        return originalFullText;
    }
    return originalFullText;
}

function processRunForReconstruction(runElement, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, nextCharCode) {
    let fullText = originalFullText;
    const rPr = getFirstElementByTag(runElement, 'w:rPr');

    Array.from(runElement.childNodes).forEach(runChild => {
        if (runChild.nodeName === 'w:t') {
            const textContent = runChild.textContent || '';
            if (textContent.length > 0) {
                propertyMap.push({
                    start: fullText.length,
                    end: fullText.length + textContent.length,
                    rPr
                });
                fullText += textContent;
            }
        } else if (runChild.nodeName === 'w:br' || runChild.nodeName === 'w:cr') {
            fullText += '\n';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (runChild.nodeName === 'w:tab') {
            fullText += '\t';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (runChild.nodeName === 'w:noBreakHyphen') {
            fullText += '\u2011';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (['w:drawing', 'w:pict', 'w:object', 'w:fldChar', 'w:instrText', 'w:sym'].includes(runChild.nodeName)) {
            const textBoxContent = getFirstElementByTag(runChild, 'w:txbxContent');
            const hasTextBox = runChild.nodeName === 'w:pict' && !!textBoxContent;

            sentinelMap.push({
                start: fullText.length,
                node: runChild,
                isTextBox: hasTextBox,
                originalContainer: hasTextBox ? textBoxContent : undefined
            });
            fullText += '\uFFFC';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (runChild.nodeName === 'w:footnoteReference' || runChild.nodeName === 'w:endnoteReference') {
            const id = runChild.getAttribute('w:id');
            if (id) {
                const type = runChild.nodeName === 'w:footnoteReference' ? 'FN' : 'EN';
                const tokenString = `{{__${type}_${id}__}}`;
                const char = String.fromCharCode(nextCharCode);
                referenceMap.set(char, runChild);
                tokenToCharMap.set(tokenString, char);
                fullText += char;
                propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
            }
        } else if (runChild.nodeName === 'w:commentReference') {
            sentinelMap.push({ start: fullText.length, node: runChild, isCommentMarker: true });
        }
    });

    return fullText;
}

function processHyperlinkForReconstruction(hyperlinkElement, originalFullText, propertyMap) {
    let fullText = originalFullText;

    Array.from(hyperlinkElement.childNodes).forEach(hyperlinkChild => {
        if (hyperlinkChild.nodeName !== 'w:r') return;

        const rPr = getFirstElementByTag(hyperlinkChild, 'w:rPr');
        const texts = getElementsByTag(hyperlinkChild, 'w:t');
        texts.forEach(textNode => {
            const textContent = textNode.textContent || '';
            if (textContent.length === 0) return;

            propertyMap.push({
                start: fullText.length,
                end: fullText.length + textContent.length,
                rPr,
                wrapper: hyperlinkElement
            });
            fullText += textContent;
        });
    });

    return fullText;
}
