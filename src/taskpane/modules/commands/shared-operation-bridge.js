/**
 * Shared operation bridge for Word add-in command handlers.
 *
 * Converts paragraph-level OOXML into a standalone-compatible document scope,
 * executes the shared operation runner, and converts the result back into
 * paragraph/package OOXML suitable for Word insertion.
 */

import { createParser, createSerializer } from '../reconciliation/adapters/xml-adapter.js';
import {
    enforceListBindingOnParagraphNodes,
    extractReplacementNodesFromOoxml,
    getParagraphListInfo,
    normalizeBodySectionOrderStandalone
} from '../reconciliation/standalone.js';
import { applyOperationToDocumentXml } from '../reconciliation/services/standalone-operation-runner.js';
import { wrapParagraphWithComments } from '../reconciliation/services/comment-package.js';
import {
    buildDocumentCommentsPackage,
    buildDocumentFragmentPackage,
    buildParagraphOnlyPackage
} from '../reconciliation/services/package-builder.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parseXmlStrict(xmlText, label) {
    const parser = createParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        throw new Error(`[XML parse error] ${label}: ${parseError.textContent || 'Unknown parse error'}`);
    }
    return xmlDoc;
}

function wrapParagraphNodesAsDocument(paragraphNodes) {
    const serializer = createSerializer();
    const bodyXml = (paragraphNodes || [])
        .map(node => serializer.serializeToString(node))
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>${bodyXml}<w:sectPr/></w:body>
</w:document>`;
}

function extractParagraphNodesFromOoxml(oxml) {
    const extracted = extractReplacementNodesFromOoxml(oxml);
    const paragraphs = (extracted.replacementNodes || [])
        .filter(node => node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'p');
    return paragraphs;
}

function extractBodyChildElements(xmlDoc) {
    if (!xmlDoc) return [];
    const body = xmlDoc.getElementsByTagNameNS(NS_W, 'body')[0] || xmlDoc.getElementsByTagNameNS('*', 'body')[0];
    if (!body) return [];

    return Array.from(body.childNodes || []).filter(
        node => node
            && node.nodeType === 1
            && !(node.namespaceURI === NS_W && node.localName === 'sectPr')
    );
}

/**
 * Applies a shared standalone operation against paragraph OOXML.
 *
 * @param {string} paragraphOoxml - OOXML from Word paragraph/range getOoxml()
 * @param {Object} operation - Canonical operation (`redline`/`highlight`/`comment`)
 * @param {Object} [options={}]
 * @param {string} [options.author='Gemini AI']
 * @param {boolean} [options.generateRedlines=true]
 * @param {(message: string) => void} [options.onInfo]
 * @param {(message: string) => void} [options.onWarn]
 * @returns {Promise<{
 *   hasChanges: boolean,
 *   paragraphOoxml?: string,
 *   packageOoxml?: string|null,
 *   commentsXml?: string|null,
 *   numberingXml?: string|null,
 *   warnings?: string[]
 * }>}
 */
export async function applySharedOperationToParagraphOoxml(paragraphOoxml, operation, options = {}) {
    const paragraphNodes = extractParagraphNodesFromOoxml(paragraphOoxml);
    if (!paragraphNodes || paragraphNodes.length === 0) {
        throw new Error('No paragraph nodes found in paragraph OOXML');
    }
    const sourceListInfo = getParagraphListInfo(paragraphNodes[0]);

    const inputDocumentXml = wrapParagraphNodesAsDocument(paragraphNodes);
    const result = await applyOperationToDocumentXml(
        inputDocumentXml,
        operation,
        options.author || 'Gemini AI',
        null,
        {
            generateRedlines: options.generateRedlines !== false,
            onInfo: options.onInfo,
            onWarn: options.onWarn
        }
    );

    if (!result?.hasChanges) {
        return {
            hasChanges: false,
            warnings: result?.warnings || []
        };
    }

    const outputDoc = parseXmlStrict(result.documentXml, 'shared operation output');
    normalizeBodySectionOrderStandalone(outputDoc);
    const outputParagraphs = Array.from(outputDoc.getElementsByTagNameNS(NS_W, 'p'));
    if (!outputParagraphs || outputParagraphs.length === 0) {
        throw new Error('Shared operation output has no paragraphs');
    }
    if (operation?.type === 'redline' && sourceListInfo?.numId) {
        enforceListBindingOnParagraphNodes([outputParagraphs[0]], {
            numId: sourceListInfo.numId,
            ilvl: sourceListInfo.ilvl || 0,
            clearParagraphPropertyChanges: true,
            removeListPropertyNode: true
        });
    }

    const serializer = createSerializer();
    const paragraphXml = serializer.serializeToString(outputParagraphs[0]);
    const commentsXml = result.commentsXml || null;
    const packageOoxml = commentsXml
        ? wrapParagraphWithComments(paragraphXml, commentsXml)
        : buildParagraphOnlyPackage(paragraphXml);

    return {
        hasChanges: true,
        paragraphOoxml: paragraphXml,
        packageOoxml,
        commentsXml,
        numberingXml: result.numberingXml || null,
        warnings: result.warnings || []
    };
}

/**
 * Applies a shared standalone operation against OOXML scope (one or more paragraphs).
 *
 * @param {string} scopeOoxml - OOXML from Word paragraph/range getOoxml()
 * @param {Object} operation - Canonical operation (`redline`/`highlight`/`comment`)
 * @param {Object} [options={}]
 * @param {string} [options.author='Gemini AI']
 * @param {boolean} [options.generateRedlines=true]
 * @param {(message: string) => void} [options.onInfo]
 * @param {(message: string) => void} [options.onWarn]
 * @returns {Promise<{
 *   hasChanges: boolean,
 *   packageOoxml?: string|null,
 *   commentsXml?: string|null,
 *   numberingXml?: string|null,
 *   warnings?: string[]
 * }>}
 */
export async function applySharedOperationToScopeOoxml(scopeOoxml, operation, options = {}) {
    const paragraphNodes = extractParagraphNodesFromOoxml(scopeOoxml);
    if (!paragraphNodes || paragraphNodes.length === 0) {
        throw new Error('No paragraph nodes found in scope OOXML');
    }

    const inputDocumentXml = wrapParagraphNodesAsDocument(paragraphNodes);
    const result = await applyOperationToDocumentXml(
        inputDocumentXml,
        operation,
        options.author || 'Gemini AI',
        null,
        {
            generateRedlines: options.generateRedlines !== false,
            onInfo: options.onInfo,
            onWarn: options.onWarn
        }
    );

    if (!result?.hasChanges) {
        return {
            hasChanges: false,
            warnings: result?.warnings || []
        };
    }

    const outputDoc = parseXmlStrict(result.documentXml, 'shared operation output');
    normalizeBodySectionOrderStandalone(outputDoc);

    const serializer = createSerializer();
    const outputBodyElements = extractBodyChildElements(outputDoc);
    const scopeXml = outputBodyElements
        .map(element => serializer.serializeToString(element))
        .join('');
    if (!scopeXml) {
        throw new Error('Shared operation output has no body elements');
    }

    const commentsXml = result.commentsXml || null;
    const numberingXml = result.numberingXml || null;

    const packageOoxml = commentsXml
        ? buildDocumentCommentsPackage(serializer.serializeToString(outputDoc.documentElement), commentsXml)
        : buildDocumentFragmentPackage(scopeXml, {
            includeNumbering: !!numberingXml,
            numberingXml,
            appendTrailingParagraph: true
        });

    return {
        hasChanges: true,
        packageOoxml,
        commentsXml,
        numberingXml,
        warnings: result.warnings || []
    };
}
