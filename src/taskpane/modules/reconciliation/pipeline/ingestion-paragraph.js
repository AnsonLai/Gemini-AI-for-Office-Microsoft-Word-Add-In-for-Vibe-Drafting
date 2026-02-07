/**
 * OOXML Reconciliation Pipeline - Paragraph Ingestion
 *
 * Parses paragraph content into run models and accepted text offsets.
 */

import { NS_W, RunKind, ContainerKind } from '../core/types.js';
import { appendParagraphBoundary, advanceOffsetForParagraphBoundary } from '../core/paragraph-offset-policy.js';
import { createParser, serializeXml } from '../adapters/xml-adapter.js';
import { warn, error as logError } from '../adapters/logger.js';
import {
    getElementsByTagNS,
    getFirstElementByTagNS,
    getFirstElementByTagNSOrTag,
    getXmlParseError
} from '../core/xml-query.js';
import { childNodesToArray, isNamespacedNode, serializeAttributes } from './ingestion-xml.js';

let containerIdCounter = 0;

/**
 * Ingests OOXML paragraph content and builds a run-aware text model.
 *
 * @param {string} ooxmlString - OOXML input
 * @returns {import('../core/types.js').IngestionResult}
 */
export function ingestOoxml(ooxmlString) {
    const runModel = [];
    let acceptedText = '';

    if (!ooxmlString) {
        return { runModel, acceptedText, pPr: null };
    }

    try {
        const parser = createParser();
        const doc = parser.parseFromString(ooxmlString, 'application/xml');

        const parseError = getXmlParseError(doc);
        if (parseError) {
            logError('OOXML parse error:', parseError.textContent);
            return { runModel, acceptedText, pPr: null };
        }

        const paragraphs = getElementsByTagNS(doc, NS_W, 'p');
        if (paragraphs.length === 0) {
            warn('No paragraphs found in OOXML');
            return { runModel, acceptedText, pPr: null };
        }

        return ingestParagraphElements(paragraphs, { includeParagraphBoundaries: true });
    } catch (error) {
        logError('Error ingesting OOXML:', error);
        return { runModel, acceptedText, pPr: null };
    }
}

/**
 * Ingests a single paragraph element directly, avoiding serialize+reparse cycles.
 *
 * @param {Element} paragraphElement - Paragraph element
 * @returns {import('../core/types.js').IngestionResult}
 */
export function ingestParagraphElement(paragraphElement) {
    if (!paragraphElement) {
        return { runModel: [], acceptedText: '', pPr: null };
    }
    return ingestParagraphElements([paragraphElement], { includeParagraphBoundaries: false });
}

/**
 * Extracts numbering metadata from a paragraph.
 *
 * @param {Element} pElement - w:p element
 * @returns {Object|null}
 */
export function detectNumberingContext(pElement) {
    const pPr = getFirstElementByTagNS(pElement, NS_W, 'pPr');
    if (!pPr) return null;

    const numPr = getFirstElementByTagNS(pPr, NS_W, 'numPr');
    if (!numPr) return null;

    const numIdEl = getFirstElementByTagNS(numPr, NS_W, 'numId');
    const ilvlEl = getFirstElementByTagNS(numPr, NS_W, 'ilvl');
    if (!numIdEl) return null;

    const numId = numIdEl.getAttribute('w:val');
    const type = numId === '1' ? 'bullet' : (numId === '2' ? 'numbered' : 'unknown');

    return {
        numId,
        ilvl: parseInt(ilvlEl?.getAttribute('w:val') || '0', 10),
        type
    };
}

function ingestParagraphElements(paragraphs, options = {}) {
    const includeParagraphBoundaries = options.includeParagraphBoundaries ?? true;
    const runModel = [];
    let acceptedText = '';
    let currentOffset = 0;
    let firstPPr = null;

    for (let pIndex = 0; pIndex < paragraphs.length; pIndex++) {
        const pElement = paragraphs[pIndex];
        const pPr = getFirstElementByTagNS(pElement, NS_W, 'pPr');
        const pPrXml = pPr ? serializeXml(pPr) : '';

        if (pIndex === 0) {
            firstPPr = pPr;
        }

        runModel.push({
            kind: RunKind.PARAGRAPH_START,
            pPrXml,
            startOffset: currentOffset,
            endOffset: currentOffset,
            text: ''
        });

        const result = processNodeRecursive(pElement, currentOffset, runModel);
        acceptedText += result.text;
        currentOffset = acceptedText.length;

        if (includeParagraphBoundaries) {
            acceptedText = appendParagraphBoundary(acceptedText, pIndex, paragraphs.length);
            currentOffset = advanceOffsetForParagraphBoundary(currentOffset, pIndex, paragraphs.length);
        }
    }

    return { runModel, acceptedText, pPr: firstPPr };
}

function processNodeRecursive(node, currentOffset, runModel) {
    let localOffset = currentOffset;
    let text = '';

    const handlers = getNodeHandlers(runModel);
    for (const child of childNodesToArray(node)) {
        if (isNamespacedNode(child, NS_W, 'pPr') || isNamespacedNode(child, NS_W, 'proofErr')) {
            continue;
        }

        const handler = handlers.get(child.localName);
        if (!handler) continue;

        const result = handler(child, localOffset);
        localOffset = result.offset;
        text += result.text;
    }

    return { offset: localOffset, text };
}

function getNodeHandlers(runModel) {
    /** @type {Map<string, (child: Node, offset: number) => { offset: number, text: string }>} */
    const handlers = new Map();

    handlers.set('sdt', (child, offset) => {
        const containerId = `sdt_${containerIdCounter++}`;
        const sdtPr = getFirstElementByTagNS(child, NS_W, 'sdtPr');
        const sdtContent = getFirstElementByTagNS(child, NS_W, 'sdtContent');

        runModel.push({
            kind: RunKind.CONTAINER_START,
            containerKind: ContainerKind.SDT,
            containerId,
            propertiesXml: sdtPr ? serializeXml(sdtPr) : '',
            startOffset: offset,
            endOffset: offset,
            text: ''
        });

        const contentResult = sdtContent
            ? processNodeRecursive(sdtContent, offset, runModel)
            : { offset, text: '' };

        runModel.push({
            kind: RunKind.CONTAINER_END,
            containerKind: ContainerKind.SDT,
            containerId,
            startOffset: contentResult.offset,
            endOffset: contentResult.offset,
            text: ''
        });

        return contentResult;
    });

    handlers.set('smartTag', (child, offset) => {
        const containerId = `smartTag_${containerIdCounter++}`;
        runModel.push({
            kind: RunKind.CONTAINER_START,
            containerKind: ContainerKind.SMART_TAG,
            containerId,
            propertiesXml: serializeAttributes(child),
            startOffset: offset,
            endOffset: offset,
            text: ''
        });

        const contentResult = processNodeRecursive(child, offset, runModel);
        runModel.push({
            kind: RunKind.CONTAINER_END,
            containerKind: ContainerKind.SMART_TAG,
            containerId,
            startOffset: contentResult.offset,
            endOffset: contentResult.offset,
            text: ''
        });

        return contentResult;
    });

    handlers.set('del', (child, offset) => {
        const deletionEntry = processDeletion(child, offset);
        if (deletionEntry) {
            runModel.push(deletionEntry);
        }
        return { offset, text: '' };
    });

    handlers.set('bookmarkStart', (child, offset) => {
        runModel.push({
            kind: RunKind.BOOKMARK,
            nodeXml: serializeXml(child),
            startOffset: offset,
            endOffset: offset,
            text: ''
        });
        return { offset, text: '' };
    });

    handlers.set('bookmarkEnd', (child, offset) => {
        runModel.push({
            kind: RunKind.BOOKMARK,
            nodeXml: serializeXml(child),
            startOffset: offset,
            endOffset: offset,
            text: ''
        });
        return { offset, text: '' };
    });

    handlers.set('ins', (child, offset) => processNodeRecursive(child, offset, runModel));

    handlers.set('hyperlink', (child, offset) => {
        const containerId = `hyperlink_${containerIdCounter++}`;
        const rId = child.getAttribute('r:id') || '';
        const anchor = child.getAttribute('w:anchor') || '';

        runModel.push({
            kind: RunKind.CONTAINER_START,
            containerKind: 'hyperlink',
            containerId,
            propertiesXml: JSON.stringify({ rId, anchor }),
            startOffset: offset,
            endOffset: offset,
            text: ''
        });

        const contentResult = processNodeRecursive(child, offset, runModel);
        runModel.push({
            kind: RunKind.CONTAINER_END,
            containerKind: 'hyperlink',
            containerId,
            startOffset: contentResult.offset,
            endOffset: contentResult.offset,
            text: ''
        });

        return contentResult;
    });

    handlers.set('r', (child, offset) => {
        const runEntry = processRun(child, offset);
        if (!runEntry || !runEntry.text) {
            return { offset, text: '' };
        }

        runModel.push(runEntry);
        const nextOffset = offset + runEntry.text.length;
        return { offset: nextOffset, text: runEntry.text };
    });

    return handlers;
}

function processRun(runElement, startOffset) {
    const rPr = getFirstElementByTagNSOrTag(runElement, NS_W, 'rPr');
    const rPrXml = rPr ? serializeXml(rPr) : '';

    let text = '';
    for (const child of childNodesToArray(runElement)) {
        const nodeName = child.nodeName;
        if (nodeName.endsWith(':t') || nodeName === 't') {
            text += child.textContent || '';
        } else if (nodeName.endsWith(':br') || nodeName === 'br' || nodeName.endsWith(':cr') || nodeName === 'cr') {
            text += '\n';
        } else if (nodeName.endsWith(':tab') || nodeName === 'tab') {
            text += '\t';
        } else if (nodeName.endsWith(':noBreakHyphen') || nodeName === 'noBreakHyphen') {
            text += '\u2011';
        }
    }

    if (!text) return null;

    return {
        kind: RunKind.TEXT,
        text,
        rPrXml,
        startOffset,
        endOffset: startOffset + text.length
    };
}

function processDeletion(delElement, offset) {
    const author = delElement.getAttribute('w:author') || '';

    let text = '';
    const delTexts = getElementsByTagNS(delElement, NS_W, 'delText');
    for (const delText of delTexts) {
        text += delText.textContent || '';
    }

    const runs = getElementsByTagNS(delElement, NS_W, 'r');
    for (const run of runs) {
        const innerDelTexts = getElementsByTagNS(run, NS_W, 'delText');
        for (const delText of innerDelTexts) {
            text += delText.textContent || '';
        }
    }

    if (!text) return null;

    return {
        kind: RunKind.DELETION,
        text,
        rPrXml: '',
        startOffset: offset,
        endOffset: offset,
        author,
        nodeXml: serializeXml(delElement)
    };
}
