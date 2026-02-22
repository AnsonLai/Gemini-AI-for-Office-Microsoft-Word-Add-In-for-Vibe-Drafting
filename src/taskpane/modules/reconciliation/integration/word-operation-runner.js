/**
 * Word adapter for applying shared standalone operations to live Word scopes.
 *
 * Bridges Word paragraph/range OOXML to the standalone operation runner, then
 * writes package OOXML back using Word integration helpers.
 */

import { createParser, createSerializer } from '../adapters/xml-adapter.js';
import {
    enforceListBindingOnParagraphNodes,
    getParagraphText,
    extractReplacementNodesFromOoxml,
    normalizeBodySectionOrderStandalone
} from '../standalone.js';
import { applyOperationToDocumentXml } from '../services/standalone-operation-runner.js';
import { wrapParagraphWithComments } from '../services/comment-package.js';
import {
    buildDocumentCommentsPackage,
    buildDocumentFragmentPackage,
    buildParagraphOnlyPackage
} from '../services/package-builder.js';
import {
    insertOoxmlWithRangeFallback,
    withNativeTrackingDisabled
} from './word-ooxml.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const SIMPLE_LIST_MARKER_RE = /^\s*(?:[-*+]\s+|\d+(?:\.\d+)*[.)]\s+|[A-Za-z][.)]\s+)/;

function getDirectWordChild(element, localName) {
    if (!element) return null;
    return Array.from(element.childNodes || []).find(
        node => node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === localName
    ) || null;
}

function readValAttribute(element) {
    if (!element || typeof element.getAttribute !== 'function') return null;
    return element.getAttribute('w:val') || element.getAttribute('val') || null;
}

function getDirectParagraphListInfo(paragraph) {
    if (!paragraph) return null;
    const pPr = getDirectWordChild(paragraph, 'pPr');
    if (!pPr) return null;
    const numPr = getDirectWordChild(pPr, 'numPr');
    if (!numPr) return null;
    const numIdEl = getDirectWordChild(numPr, 'numId');
    if (!numIdEl) return null;
    const numId = readValAttribute(numIdEl);
    if (!numId) return null;

    const ilvlEl = getDirectWordChild(numPr, 'ilvl');
    const ilvlRaw = readValAttribute(ilvlEl);
    const ilvl = Number.parseInt(ilvlRaw || '0', 10);
    return {
        numId: String(numId),
        ilvl: Number.isFinite(ilvl) ? ilvl : 0
    };
}

function isSimplePlainTextRedline(operation) {
    if (operation?.type !== 'redline') return false;
    const modified = String(operation?.modified || '');
    if (!modified.trim()) return false;
    if (modified.includes('\n')) return false;
    if (modified.includes('|') && modified.includes('---')) return false;
    if (SIMPLE_LIST_MARKER_RE.test(modified)) return false;
    return true;
}

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
    return (extracted.replacementNodes || [])
        .filter(node => node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'p');
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

function normalizeTextForParagraphSelection(text) {
    return String(text || '').replace(/\s+/g, ' ').trim();
}

function selectParagraphNodesForParagraphScope(paragraphNodes, operation) {
    if (!Array.isArray(paragraphNodes) || paragraphNodes.length === 0) return [];
    if (paragraphNodes.length === 1) return [paragraphNodes[0]];

    const normalizedTarget = normalizeTextForParagraphSelection(operation?.target);
    if (normalizedTarget) {
        const exactMatch = paragraphNodes.find(node =>
            normalizeTextForParagraphSelection(getParagraphText(node)) === normalizedTarget
        );
        if (exactMatch) return [exactMatch];
    }

    const firstNonEmpty = paragraphNodes.find(node =>
        normalizeTextForParagraphSelection(getParagraphText(node)).length > 0
    );
    return [firstNonEmpty || paragraphNodes[0]];
}

/**
 * Applies a shared standalone operation against paragraph OOXML.
 *
 * @param {string} paragraphOoxml
 * @param {Object} operation
 * @param {Object} [options={}]
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
    const scopedParagraphNodes = selectParagraphNodesForParagraphScope(paragraphNodes, operation);
    if (!scopedParagraphNodes || scopedParagraphNodes.length === 0) {
        throw new Error('Unable to isolate target paragraph node for shared operation');
    }
    const sourceDirectListInfo = getDirectParagraphListInfo(scopedParagraphNodes[0]);

    const inputDocumentXml = wrapParagraphNodesAsDocument(scopedParagraphNodes);
    const runner = typeof options.runner === 'function' ? options.runner : applyOperationToDocumentXml;
    const result = await runner(
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
            singleParagraphOutput: false,
            warnings: result?.warnings || []
        };
    }

    const outputDoc = parseXmlStrict(result.documentXml, 'shared operation output');
    normalizeBodySectionOrderStandalone(outputDoc);
    const outputParagraphs = Array.from(outputDoc.getElementsByTagNameNS(NS_W, 'p'));
    if (!outputParagraphs || outputParagraphs.length === 0) {
        throw new Error('Shared operation output has no paragraphs');
    }
    const outputBodyElements = extractBodyChildElements(outputDoc);
    const isSingleParagraphOutput =
        outputBodyElements.length === 1
        && outputBodyElements[0].namespaceURI === NS_W
        && outputBodyElements[0].localName === 'p';

    if (operation?.type === 'redline' && sourceDirectListInfo?.numId) {
        enforceListBindingOnParagraphNodes([outputParagraphs[0]], {
            numId: sourceDirectListInfo.numId,
            ilvl: sourceDirectListInfo.ilvl || 0,
            clearParagraphPropertyChanges: true,
            removeListPropertyNode: true
        });
    }

    const serializer = createSerializer();
    const paragraphXml = serializer.serializeToString(outputParagraphs[0]);
    const commentsXml = result.commentsXml || null;
    const numberingXml = result.numberingXml || null;
    // Cuando el párrafo de origen era un elemento de lista y se aplicó enforceListBinding,
    // usamos buildParagraphOnlyPackage en lugar de buildDocumentFragmentPackage.
    // El paquete de fragmento no incluye las definiciones de numeración del documento (ej. numId=6),
    // por lo que Word no puede resolver el numId original y vuelve al estilo de viñeta incorrecto.
    // El paquete de párrafo único se inserta en el contexto vivo del documento donde la numeración ya existe.
    const listItemRedlineEnforced = operation?.type === 'redline' && !!sourceDirectListInfo?.numId && !commentsXml;
    const useParagraphOnlyListPackage = listItemRedlineEnforced && isSingleParagraphOutput;
    const packageOoxml = isSingleParagraphOutput
        ? (
            commentsXml
                ? wrapParagraphWithComments(paragraphXml, commentsXml)
                : buildParagraphOnlyPackage(paragraphXml)
        )
        : useParagraphOnlyListPackage
            ? buildParagraphOnlyPackage(paragraphXml)
            : (
                commentsXml
                    ? buildDocumentCommentsPackage(serializer.serializeToString(outputDoc.documentElement), commentsXml)
                    : buildDocumentFragmentPackage(
                        outputBodyElements.map(element => serializer.serializeToString(element)).join(''),
                        {
                            includeNumbering: !!numberingXml,
                            numberingXml,
                            appendTrailingParagraph: true
                        }
                    )
            );

    return {
        hasChanges: true,
        singleParagraphOutput: isSingleParagraphOutput,
        paragraphOoxml: paragraphXml,
        packageOoxml,
        commentsXml,
        numberingXml,
        warnings: result.warnings || []
    };
}

/**
 * Applies a shared standalone operation against OOXML scope (one or more paragraphs).
 *
 * @param {string} scopeOoxml
 * @param {Object} operation
 * @param {Object} [options={}]
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
    const runner = typeof options.runner === 'function' ? options.runner : applyOperationToDocumentXml;
    const result = await runner(
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
            singleParagraphOutput: false,
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
        singleParagraphOutput: false,
        packageOoxml,
        commentsXml,
        numberingXml,
        warnings: result.warnings || []
    };
}

function resolveWordOperationScope(scope) {
    if (!scope) throw new Error('Missing scope for applyWordOperation');
    if (scope.paragraph) {
        if (scope.endParagraph) {
            const range = scope.paragraph.getRange().expandTo(scope.endParagraph.getRange());
            return { kind: 'range', target: range };
        }
        return { kind: 'paragraph', target: scope.paragraph };
    }
    if (scope.range) {
        return { kind: 'range', target: scope.range };
    }
    if (typeof scope.getOoxml === 'function' && typeof scope.insertOoxml === 'function') {
        return { kind: 'range', target: scope };
    }
    throw new Error('Unsupported scope shape for applyWordOperation');
}

/**
 * Applies a canonical operation to a Word paragraph/range scope.
 *
 * @param {Word.RequestContext} context
 * @param {Object} operation
 * @param {Object} scope
 * @param {Object} [options={}]
 * @returns {Promise<boolean>} True when changes were applied
 */
export async function applyWordOperation(context, operation, scope, options = {}) {
    const resolved = resolveWordOperationScope(scope);
    const scopeOoxmlResult = resolved.target.getOoxml();
    await context.sync();

    const bridgeOptions = {
        author: options.author,
        generateRedlines: options.generateRedlines,
        onInfo: options.onInfo,
        onWarn: options.onWarn,
        runner: options.runner
    };
    const bridgeResult = resolved.kind === 'paragraph'
        ? await applySharedOperationToParagraphOoxml(scopeOoxmlResult?.value || '', operation, bridgeOptions)
        : await applySharedOperationToScopeOoxml(scopeOoxmlResult?.value || '', operation, bridgeOptions);

    if (!bridgeResult.hasChanges) {
        return false;
    }

    const canUseDirectParagraphPayload = false; // direct_paragraph snippet injection crashes Word parser with w:ins elements
    const insertionPayload = canUseDirectParagraphPayload
        ? bridgeResult.paragraphOoxml
        : bridgeResult.packageOoxml;
    if (!insertionPayload) {
        return false;
    }
    if (typeof options.onInfo === 'function') {
        options.onInfo(
            `Insertion strategy: ${canUseDirectParagraphPayload ? 'direct_paragraph' : 'package'} `
            + `(kind=${resolved.kind}, singleParagraphOutput=${bridgeResult.singleParagraphOutput === true}, `
            + `hasComments=${!!bridgeResult.commentsXml}, hasNumbering=${!!bridgeResult.numberingXml}, `
            + `isSimplePlainTextRedline=${isSimplePlainTextRedline(operation)})`
        );
    }

    await withNativeTrackingDisabled(context, async () => {
        if (resolved.kind === 'paragraph') {
            await insertOoxmlWithRangeFallback(
                resolved.target,
                insertionPayload,
                'Replace',
                context,
                options.logPrefix || 'WordOp/Shared'
            );
        } else {
            const insertMode = (typeof Word !== 'undefined' && (Word.InsertLocation?.replace || Word.InsertLocation?.Replace))
                || 'Replace';
            resolved.target.insertOoxml(insertionPayload, insertMode);
            await context.sync();
        }
    }, {
        enabled: !!options.disableNativeTracking,
        baseTrackingMode: options.baseTrackingMode ?? null,
        logPrefix: options.logPrefix || 'WordOp/Shared'
    });

    return true;
}

/**
 * Applies a shared standalone operation to a single Word paragraph.
 *
 * @param {Object} params
 * @param {Word.RequestContext} params.context
 * @param {Word.Paragraph} params.targetParagraph
 * @param {Object} params.operation
 * @param {string} params.author
 * @param {boolean} params.generateRedlines
 * @param {boolean} [params.disableNativeTracking=false]
 * @param {Word.ChangeTrackingMode|null} [params.baseTrackingMode=null]
 * @param {string} params.logPrefix
 * @returns {Promise<boolean>} True when a change is applied
 */
export async function applySharedOperationToWordParagraph({
    context,
    targetParagraph,
    operation,
    author,
    generateRedlines,
    disableNativeTracking = false,
    baseTrackingMode = null,
    logPrefix
}) {
    return applyWordOperation(context, operation, { paragraph: targetParagraph }, {
        author,
        generateRedlines,
        disableNativeTracking,
        baseTrackingMode,
        logPrefix,
        onInfo: message => console.log(`[${logPrefix}] ${message}`),
        onWarn: message => console.warn(`[${logPrefix}] ${message}`)
    });
}

/**
 * Applies a shared standalone operation to a Word paragraph/range scope.
 *
 * @param {Object} params
 * @param {Word.RequestContext} params.context
 * @param {Word.Range|Word.Paragraph} params.scope
 * @param {Object} params.operation
 * @param {string} params.author
 * @param {boolean} params.generateRedlines
 * @param {boolean} [params.disableNativeTracking=false]
 * @param {Word.ChangeTrackingMode|null} [params.baseTrackingMode=null]
 * @param {string} params.logPrefix
 * @returns {Promise<boolean>} True when a change is applied
 */
export async function applySharedOperationToWordScope({
    context,
    scope,
    operation,
    author,
    generateRedlines,
    disableNativeTracking = false,
    baseTrackingMode = null,
    logPrefix
}) {
    return applyWordOperation(context, operation, { range: scope }, {
        author,
        generateRedlines,
        disableNativeTracking,
        baseTrackingMode,
        logPrefix,
        onInfo: message => console.log(`[${logPrefix}] ${message}`),
        onWarn: message => console.warn(`[${logPrefix}] ${message}`)
    });
}
