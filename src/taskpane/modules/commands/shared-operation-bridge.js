/**
 * Shared operation bridge for Word add-in command handlers.
 *
 * Converts paragraph-level OOXML into a standalone-compatible document scope,
 * executes the shared operation runner, and converts the result back into
 * paragraph/package OOXML suitable for Word insertion.
 */

import { createParser, createSerializer } from '../reconciliation/adapters/xml-adapter.js';
import {
    extractReplacementNodesFromOoxml,
    normalizeBodySectionOrderStandalone
} from '../reconciliation/standalone.js';
import { applyOperationToDocumentXml } from '../reconciliation/services/standalone-operation-runner.js';
import { wrapParagraphWithComments } from '../reconciliation/services/comment-package.js';
import { buildParagraphOnlyPackage } from '../reconciliation/services/package-builder.js';

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
