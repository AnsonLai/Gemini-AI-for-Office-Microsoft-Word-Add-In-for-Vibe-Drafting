/**
 * Reconstruction reconciliation mode.
 *
 * This mode rebuilds paragraph content and supports structural text changes,
 * including newline-driven paragraph/list transformations.
 */

import { diff_match_patch } from 'diff-match-patch';
import { getApplicableFormatHints } from '../pipeline/markdown-processor.js';
import { wordsToChars, charsToWords } from '../pipeline/diff-engine.js';
import { getDocumentParagraphs } from './format-extraction.js';
import { createTrackChange, createFormattedRuns } from './run-builders.js';

/**
 * Applies reconstruction mode reconciliation.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} originalText - Original text
 * @param {string} modifiedText - Modified text
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {string} author - Author name
 * @param {Array} formatHints - Format hints
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
export function applyReconstructionMode(xmlDoc, originalText, modifiedText, serializer, author, formatHints, generateRedlines = true) {
    const rootElement = xmlDoc.documentElement;
    const isBodyRoot = rootElement.nodeName === 'w:body' || rootElement.nodeName.endsWith(':package');
    const paragraphs = getDocumentParagraphs(xmlDoc);

    let body = xmlDoc.getElementsByTagName('w:body')[0];
    if (!body && isBodyRoot) body = rootElement;

    if (paragraphs.length === 0) {
        return { oxml: serializer.serializeToString(xmlDoc), hasChanges: false };
    }

    let originalFullText = '';
    const propertyMap = [];
    const paragraphMap = [];
    const sentinelMap = [];
    const referenceMap = new Map();
    const tokenToCharMap = new Map();
    let nextCharCode = 0xe000;

    const uniqueContainers = new Set();
    const replacementContainers = new Map();

    paragraphs.forEach((p, pIndex) => {
        const pStart = originalFullText.length;

        Array.from(p.childNodes).forEach(child => {
            originalFullText = processChildNode(
                child, originalFullText, propertyMap, sentinelMap,
                referenceMap, tokenToCharMap, nextCharCode
            );
            if (referenceMap.size > tokenToCharMap.size) {
                nextCharCode++;
            }
        });

        if (pIndex < paragraphs.length - 1) {
            originalFullText += '\n';
        }

        const pEnd = originalFullText.length;
        const pPr = p.getElementsByTagName('w:pPr')[0] || null;
        const container = p.parentNode;
        if (container) uniqueContainers.add(container);

        paragraphMap.push({
            start: pStart,
            end: pEnd,
            pPr,
            container: container || body
        });
    });

    let processedModifiedText = modifiedText;
    tokenToCharMap.forEach((char, tokenString) => {
        const escapedToken = tokenString.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
        processedModifiedText = processedModifiedText.replace(new RegExp(escapedToken, 'g'), char);
    });

    const dmp = new diff_match_patch();
    const { chars1, chars2, wordArray } = wordsToChars(originalFullText, processedModifiedText);
    const charDiffs = dmp.diff_main(chars1, chars2);
    dmp.diff_cleanupSemantic(charDiffs);
    const diffs = charsToWords(charDiffs, wordArray);

    const containerFragments = new Map();
    uniqueContainers.forEach(c => containerFragments.set(c, xmlDoc.createDocumentFragment()));

    if (body && !containerFragments.has(body)) {
        containerFragments.set(body, xmlDoc.createDocumentFragment());
    }

    if (!containerFragments.has(xmlDoc)) {
        containerFragments.set(xmlDoc, xmlDoc.createDocumentFragment());
    }

    const getParagraphInfo = (index) => {
        const match = paragraphMap.find(m => index >= m.start && index < m.end);
        if (!match && paragraphMap.length > 0) {
            return paragraphMap[paragraphMap.length - 1];
        }
        return match || { pPr: null, container: body };
    };

    const getRunProperties = (index) => {
        const match = propertyMap.find(m => index >= m.start && index < m.end);
        return match ? { rPr: match.rPr, wrapper: match.wrapper } : { rPr: null };
    };

    const createNewParagraph = (pPr) => {
        const newP = xmlDoc.createElement('w:p');
        if (pPr) newP.appendChild(pPr.cloneNode(true));
        return newP;
    };

    const startInfo = getParagraphInfo(0);
    let currentParagraph = createNewParagraph(startInfo.pPr);
    const currentContainer = startInfo.container;
    const currentFragment = containerFragments.get(currentContainer);
    if (currentFragment) currentFragment.appendChild(currentParagraph);

    let currentOriginalIndex = 0;
    let currentInsertOffset = 0;

    for (const [op, text] of diffs) {
        if (op === 0) {
            let offset = 0;
            while (offset < text.length) {
                const props = getRunProperties(currentOriginalIndex + offset);
                const range = propertyMap.find(m =>
                    currentOriginalIndex + offset >= m.start && currentOriginalIndex + offset < m.end
                );
                const length = range
                    ? Math.min(range.end - (currentOriginalIndex + offset), text.length - offset)
                    : 1;
                const chunk = text.substring(offset, offset + length);

                appendTextToCurrent(
                    xmlDoc, chunk, 'equal', props.rPr, props.wrapper,
                    currentOriginalIndex + offset, currentParagraph, paragraphMap,
                    containerFragments, sentinelMap, referenceMap, tokenToCharMap,
                    replacementContainers, getParagraphInfo, createNewParagraph, author,
                    formatHints, currentInsertOffset, generateRedlines
                );

                currentInsertOffset += length;
                offset += length;
            }
            currentOriginalIndex += text.length;
        } else if (op === 1) {
            const isStartOfParagraph = paragraphMap.some(p => p.start === currentOriginalIndex);
            const props = currentOriginalIndex > 0 && !isStartOfParagraph
                ? getRunProperties(currentOriginalIndex - 1)
                : getRunProperties(currentOriginalIndex);

            appendTextToCurrent(
                xmlDoc, text, 'insert', props.rPr, props.wrapper,
                currentOriginalIndex, currentParagraph, paragraphMap,
                containerFragments, sentinelMap, referenceMap, tokenToCharMap,
                replacementContainers, getParagraphInfo, createNewParagraph, author,
                formatHints, currentInsertOffset, generateRedlines
            );
            currentInsertOffset += text.length;
        } else if (op === -1) {
            let offset = 0;
            while (offset < text.length) {
                const props = getRunProperties(currentOriginalIndex + offset);
                const range = propertyMap.find(m =>
                    currentOriginalIndex + offset >= m.start && currentOriginalIndex + offset < m.end
                );
                const length = range
                    ? Math.min(range.end - (currentOriginalIndex + offset), text.length - offset)
                    : 1;
                const chunk = text.substring(offset, offset + length);

                appendTextToCurrent(
                    xmlDoc, chunk, 'delete', props.rPr, props.wrapper,
                    currentOriginalIndex + offset, currentParagraph, paragraphMap,
                    containerFragments, sentinelMap, referenceMap, tokenToCharMap,
                    replacementContainers, getParagraphInfo, createNewParagraph, author,
                    formatHints, currentInsertOffset, generateRedlines
                );

                offset += length;
            }
            currentOriginalIndex += text.length;
        }
    }

    paragraphs.forEach(p => {
        if (p.parentNode) p.parentNode.removeChild(p);
    });

    containerFragments.forEach((fragment, container) => {
        const replacement = replacementContainers.get(container);
        const target = replacement || container;

        if (target.nodeType === 9) {
            const firstChild = fragment.firstChild;
            if (firstChild) {
                target.appendChild(firstChild);
                while (fragment.firstChild) {
                    target.documentElement.appendChild(fragment.firstChild);
                }
            }
        } else {
            target.appendChild(fragment);
        }
    });

    return { oxml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

/**
 * Processes a child node during paragraph traversal.
 *
 * @param {Node} child - Child node
 * @param {string} originalFullText - Current full text
 * @param {Array} propertyMap - Property map
 * @param {Array} sentinelMap - Sentinel map
 * @param {Map<string, Node>} referenceMap - Reference char->node map
 * @param {Map<string, string>} tokenToCharMap - Token->char map
 * @param {number} nextCharCode - Next private-use char code
 * @returns {string}
 */
export function processChildNode(child, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, nextCharCode) {
    if (child.nodeName === 'w:r') {
        return processRunForReconstruction(child, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, nextCharCode);
    } else if (child.nodeName === 'w:hyperlink') {
        return processHyperlinkForReconstruction(child, originalFullText, propertyMap);
    } else if (['w:sdt', 'w:oMath', 'm:oMath', 'w:bookmarkStart', 'w:bookmarkEnd'].includes(child.nodeName)) {
        sentinelMap.push({ start: originalFullText.length, node: child });
        return originalFullText + '\uFFFC';
    } else if (['w:commentRangeStart', 'w:commentRangeEnd'].includes(child.nodeName)) {
        sentinelMap.push({ start: originalFullText.length, node: child, isCommentMarker: true });
        return originalFullText;
    }
    return originalFullText;
}

/**
 * Processes run text and inline run children for reconstruction mapping.
 *
 * @param {Element} r - Run element
 * @param {string} originalFullText - Current full text
 * @param {Array} propertyMap - Property map
 * @param {Array} sentinelMap - Sentinel map
 * @param {Map<string, Node>} referenceMap - Reference char->node map
 * @param {Map<string, string>} tokenToCharMap - Token->char map
 * @param {number} nextCharCode - Next private-use char code
 * @returns {string}
 */
export function processRunForReconstruction(r, originalFullText, propertyMap, sentinelMap, referenceMap, tokenToCharMap, nextCharCode) {
    let fullText = originalFullText;
    const rPr = r.getElementsByTagName('w:rPr')[0] || null;

    Array.from(r.childNodes).forEach(rc => {
        if (rc.nodeName === 'w:t') {
            const textContent = rc.textContent || '';
            if (textContent.length > 0) {
                propertyMap.push({
                    start: fullText.length,
                    end: fullText.length + textContent.length,
                    rPr
                });
                fullText += textContent;
            }
        } else if (rc.nodeName === 'w:br' || rc.nodeName === 'w:cr') {
            fullText += '\n';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (rc.nodeName === 'w:tab') {
            fullText += '\t';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (rc.nodeName === 'w:noBreakHyphen') {
            fullText += '\u2011';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (['w:drawing', 'w:pict', 'w:object', 'w:fldChar', 'w:instrText', 'w:sym'].includes(rc.nodeName)) {
            const rcElement = rc;
            const txbxContent = rcElement.getElementsByTagName ? rcElement.getElementsByTagName('w:txbxContent')[0] : null;
            const hasTextBox = rc.nodeName === 'w:pict' && !!txbxContent;

            sentinelMap.push({
                start: fullText.length,
                node: rc,
                isTextBox: hasTextBox,
                originalContainer: hasTextBox ? txbxContent : undefined
            });
            fullText += '\uFFFC';
            propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
        } else if (rc.nodeName === 'w:footnoteReference' || rc.nodeName === 'w:endnoteReference') {
            const ref = rc;
            const id = ref.getAttribute('w:id');
            if (id) {
                const type = rc.nodeName === 'w:footnoteReference' ? 'FN' : 'EN';
                const tokenString = `{{__${type}_${id}__}}`;
                const char = String.fromCharCode(nextCharCode);
                referenceMap.set(char, rc);
                tokenToCharMap.set(tokenString, char);
                fullText += char;
                propertyMap.push({ start: fullText.length - 1, end: fullText.length, rPr });
            }
        } else if (rc.nodeName === 'w:commentReference') {
            sentinelMap.push({ start: fullText.length, node: rc, isCommentMarker: true });
        }
    });

    return fullText;
}

/**
 * Processes hyperlink text runs for reconstruction mapping.
 *
 * @param {Element} h - Hyperlink element
 * @param {string} originalFullText - Current full text
 * @param {Array} propertyMap - Property map
 * @returns {string}
 */
export function processHyperlinkForReconstruction(h, originalFullText, propertyMap) {
    let fullText = originalFullText;

    Array.from(h.childNodes).forEach(hc => {
        if (hc.nodeName === 'w:r') {
            const r = hc;
            const rPr = r.getElementsByTagName('w:rPr')[0] || null;
            const texts = Array.from(r.getElementsByTagName('w:t'));
            texts.forEach(t => {
                const textContent = t.textContent || '';
                if (textContent.length > 0) {
                    propertyMap.push({
                        start: fullText.length,
                        end: fullText.length + textContent.length,
                        rPr,
                        wrapper: h
                    });
                    fullText += textContent;
                }
            });
        }
    });

    return fullText;
}

/**
 * Appends text segments to the active reconstruction paragraph.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} text - Text chunk
 * @param {'equal'|'insert'|'delete'} type - Diff op type
 * @param {Element|null} rPr - Run properties
 * @param {Element|null} wrapper - Optional wrapper
 * @param {number} baseIndex - Base original index
 * @param {Element} currentParagraphRef - Active paragraph
 * @param {Array} paragraphMap - Paragraph map
 * @param {Map<Node, DocumentFragment>} containerFragments - Container fragment map
 * @param {Array} sentinelMap - Sentinel map
 * @param {Map<string, Node>} referenceMap - Reference map
 * @param {Map<string, string>} tokenToCharMap - Token->char map
 * @param {Map<Node, Node>} replacementContainers - Replacement container map
 * @param {Function} getParagraphInfo - Paragraph resolver
 * @param {Function} createNewParagraph - Paragraph factory
 * @param {string} author - Author name
 * @param {Array} [formatHints=[]] - Format hints
 * @param {number} [insertOffset=0] - Insert offset
 * @param {boolean} [generateRedlines=true] - Track change toggle
 */
export function appendTextToCurrent(
    xmlDoc, text, type, rPr, wrapper, baseIndex,
    currentParagraphRef, paragraphMap, containerFragments,
    sentinelMap, referenceMap, tokenToCharMap,
    replacementContainers, getParagraphInfo, createNewParagraph, author,
    formatHints = [], insertOffset = 0, generateRedlines = true
) {
    let localBaseIndex = baseIndex;
    let localInsertOffset = insertOffset;
    let localParagraph = currentParagraphRef;

    const parts = text.split(/([\n\uFFFC]|[\uE000-\uF8FF])/);

    parts.forEach(part => {
        const markers = sentinelMap.filter(s => s.isCommentMarker && s.start === localBaseIndex);
        markers.forEach(marker => {
            if (marker.node.nodeName === 'w:commentReference') {
                const run = xmlDoc.createElement('w:r');
                run.appendChild(marker.node.cloneNode(true));
                localParagraph.appendChild(run);
            } else {
                localParagraph.appendChild(marker.node.cloneNode(true));
            }
        });

        if (part === '\n') {
            if (type !== 'delete') {
                const info = getParagraphInfo(localBaseIndex);
                const nextParagraph = createNewParagraph(info.pPr);
                const fragment = containerFragments.get(info.container);
                if (fragment) {
                    fragment.appendChild(nextParagraph);
                    localParagraph = nextParagraph;
                }
            }
            localBaseIndex++;
            if (type !== 'delete') {
                localInsertOffset++;
            }
        } else if (part === '\uFFFC') {
            const sentinel = sentinelMap.find(s => s.start === localBaseIndex);
            if (sentinel) {
                const clone = sentinel.node.cloneNode(true);
                if (sentinel.isTextBox && sentinel.originalContainer) {
                    const newContainer = clone.getElementsByTagName('w:txbxContent')[0];
                    if (newContainer) {
                        while (newContainer.firstChild) newContainer.removeChild(newContainer.firstChild);
                        replacementContainers.set(sentinel.originalContainer, newContainer);
                    }
                }
                localParagraph.appendChild(clone);
            }
            localBaseIndex++;
            if (type !== 'delete') {
                localInsertOffset++;
            }
        } else if (referenceMap.has(part)) {
            if (type !== 'delete') {
                const refNode = referenceMap.get(part);
                if (refNode) {
                    const clone = refNode.cloneNode(true);
                    const run = xmlDoc.createElement('w:r');
                    if (rPr) run.appendChild(rPr.cloneNode(true));
                    run.appendChild(clone);
                    localParagraph.appendChild(run);
                }
            }
            localBaseIndex++;
            if (type !== 'delete') {
                localInsertOffset++;
            }
        } else if (part.length > 0) {
            let parent = localParagraph;
            if (wrapper) {
                const wrapperClone = wrapper.cloneNode(false);
                parent = wrapperClone;
                localParagraph.appendChild(wrapperClone);
            }

            if (type === 'delete') {
                const run = xmlDoc.createElement('w:r');
                if (rPr) run.appendChild(rPr.cloneNode(true));
                const t = xmlDoc.createElement('w:delText');
                t.setAttribute('xml:space', 'preserve');
                t.textContent = part;
                run.appendChild(t);

                if (generateRedlines) {
                    const del = createTrackChange(xmlDoc, 'del', run, author);
                    parent.appendChild(del);
                }
            } else if (type === 'insert') {
                const applicableHints = getApplicableFormatHints(formatHints, localInsertOffset, localInsertOffset + part.length);
                const runs = createFormattedRuns(xmlDoc, part, rPr, applicableHints, localInsertOffset, author, generateRedlines);

                if (generateRedlines) {
                    const ins = createTrackChange(xmlDoc, 'ins', null, author);
                    runs.forEach(run => ins.appendChild(run));
                    parent.appendChild(ins);
                } else {
                    runs.forEach(run => parent.appendChild(run));
                }
            } else {
                const applicableHints = getApplicableFormatHints(formatHints, localInsertOffset, localInsertOffset + part.length);
                const runs = createFormattedRuns(xmlDoc, part, rPr, applicableHints, localInsertOffset, author, generateRedlines);
                runs.forEach(run => parent.appendChild(run));
            }
            if (type !== 'delete') {
                localInsertOffset += part.length;
            }
            localBaseIndex += part.length;
        }
    });
}
