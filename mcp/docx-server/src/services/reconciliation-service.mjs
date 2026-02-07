import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
    applyRedlineToOxml,
    configureLogger,
    configureXmlProvider,
    ingestOoxml,
    injectCommentsIntoOoxml
} from '../../../../src/taskpane/modules/reconciliation/standalone.js';
import { getLocalName, isElement, parseXmlStrict, toArray, allByTag } from './xml-utils.mjs';

configureXmlProvider({
    DOMParser,
    XMLSerializer
});

// Keep MCP output clean unless debugging is needed.
configureLogger({
    log: () => {},
    warn: () => {},
    error: () => {}
});

/**
 * @param {{
 *   paragraphXml: string,
 *   paragraphText: string,
 *   modifiedText: string,
 *   paraId?: string|null,
 *   author?: string,
 *   generateRedlines?: boolean
 * }} input
 */
export async function reconcileParagraphEdit(input) {
    const author = input.author || 'MCP AI';
    const generateRedlines = input.generateRedlines ?? true;

    const result = await applyRedlineToOxml(
        input.paragraphXml,
        input.paragraphText,
        input.modifiedText,
        {
            author,
            generateRedlines,
            targetParagraphId: input.paraId || null
        }
    );

    if (!result || result.useNativeApi) {
        throw new Error('Edit requires Word native API fallback and cannot be completed in local MCP mode');
    }

    if (!result.hasChanges) {
        return {
            hasChanges: false,
            replacementNodes: [],
            numberingXml: null,
            sourceType: 'none'
        };
    }

    if (!result.oxml) {
        throw new Error('Reconciliation returned no OOXML output');
    }

    const extracted = extractReplacementNodes(result.oxml);
    if (extracted.replacementNodes.length === 0) {
        throw new Error('Reconciliation output had no replacement nodes');
    }

    return {
        hasChanges: true,
        replacementNodes: extracted.replacementNodes,
        numberingXml: extracted.numberingXml,
        sourceType: extracted.sourceType
    };
}

/**
 * @param {{
 *   documentXml: string,
 *   paragraphIndex: number,
 *   textToFind: string,
 *   commentContent: string,
 *   author?: string
 * }} input
 */
export function reconcileAddComment(input) {
    const author = input.author || 'MCP AI';
    return injectCommentsIntoOoxml(
        input.documentXml,
        [{
            paragraphIndex: input.paragraphIndex,
            textToFind: input.textToFind,
            commentContent: input.commentContent
        }],
        { author }
    );
}

/**
 * @param {string} paragraphXml
 * @returns {string}
 */
export function deriveParagraphAcceptedText(paragraphXml) {
    const ingestion = ingestOoxml(paragraphXml);
    return ingestion.acceptedText || '';
}

/**
 * @param {string} oxml
 */
function extractReplacementNodes(oxml) {
    if (oxml.includes('<pkg:package')) {
        const pkgDoc = parseXmlStrict(oxml, 'pkg:package');
        const partNodes = allByTag(pkgDoc, 'part');

        const documentPart = partNodes.find(part => {
            const name = part.getAttribute('pkg:name') || part.getAttribute('name') || '';
            return name === '/word/document.xml';
        });
        if (!documentPart) {
            throw new Error('Package output missing /word/document.xml part');
        }

        const xmlData = allByTag(documentPart, 'xmlData')[0];
        const payloadNode = firstElementChild(xmlData);
        if (!payloadNode) {
            throw new Error('Document part is missing XML payload');
        }

        const replacements = extractBodyChildren(payloadNode);
        trimTrailingInsertionBlankParagraph(replacements);

        const numberingPart = partNodes.find(part => {
            const name = part.getAttribute('pkg:name') || part.getAttribute('name') || '';
            return name === '/word/numbering.xml';
        });
        const numberingXml = numberingPart
            ? serializeFirstElementChild(numberingPart)
            : null;

        return { replacementNodes: replacements, numberingXml, sourceType: 'package' };
    }

    if (oxml.includes('<w:document')) {
        const doc = parseXmlStrict(oxml, 'w:document');
        const replacements = extractBodyChildren(doc.documentElement);
        return { replacementNodes: replacements, numberingXml: null, sourceType: 'document' };
    }

    const wrapped = `<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${oxml}</root>`;
    const fragDoc = parseXmlStrict(wrapped, 'fragment');
    const replacementNodes = toArray(fragDoc.documentElement.childNodes).filter(node => isElement(node));
    return { replacementNodes, numberingXml: null, sourceType: 'fragment' };
}

/**
 * @param {Element} node
 * @returns {Element[]}
 */
function extractBodyChildren(node) {
    const body = allByTag(node, 'body')[0];
    if (!body) {
        return [node];
    }

    return toArray(body.childNodes).filter(child => {
        return isElement(child) && getLocalName(child) !== 'sectPr';
    });
}

/**
 * Drop the Word insertion shim paragraph (`<w:p><w:pPr/></w:p>`) that list packaging adds.
 *
 * @param {Element[]} nodes
 */
function trimTrailingInsertionBlankParagraph(nodes) {
    if (nodes.length <= 1) return;
    const last = nodes[nodes.length - 1];
    if (!last || getLocalName(last) !== 'p') return;

    const textNodes = allByTag(last, 't');
    const hasText = textNodes.some(node => (node.textContent || '').trim().length > 0);
    const hasNonPPrElement = toArray(last.childNodes).some(child => {
        return isElement(child) && getLocalName(child) !== 'pPr';
    });

    if (!hasText && !hasNonPPrElement) {
        nodes.pop();
    }
}

/**
 * @param {Element|null|undefined} parent
 * @returns {Element|null}
 */
function firstElementChild(parent) {
    if (!parent) return null;
    for (const child of toArray(parent.childNodes)) {
        if (isElement(child)) return child;
    }
    return null;
}

/**
 * @param {Element} part
 * @returns {string|null}
 */
function serializeFirstElementChild(part) {
    const xmlData = allByTag(part, 'xmlData')[0];
    const payload = firstElementChild(xmlData);
    if (!payload) return null;
    return new XMLSerializer().serializeToString(payload);
}

