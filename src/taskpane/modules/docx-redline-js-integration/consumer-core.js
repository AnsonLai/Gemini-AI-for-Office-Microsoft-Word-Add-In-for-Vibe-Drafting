/**
 * Portable OOXML consumer core: source preparation, operation mapping and result handling.
 * No Word, UI, or filesystem APIs are used in this module.
 */

import {
    enforceListBindingOnParagraphNodes,
    getParagraphText,
    extractReplacementNodesFromOoxml,
    normalizeBodySectionOrderStandalone,
    getDefaultAuthor,
    acceptTrackedChangesInOoxml,
    inspectDocumentParts,
    mergeNumberingXmlBySchemaOrder
} from '@ansonlai/docx-redline-js';
import {
    createParser,
    createSerializer
} from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';
import {
    applyOperationToDocumentXml,
    applyOperationsToDocumentXml
} from '@ansonlai/docx-redline-js/services/standalone-operation-runner.js';
import { wrapParagraphWithComments } from '@ansonlai/docx-redline-js/services/comment-package.js';
import { reconcileCommentSiblingParts } from '@ansonlai/docx-redline-js/services/comment-thread-parts.js';
import { getPartSpec } from '@ansonlai/docx-redline-js/services/package-parts.js';
import {
    buildDocumentCommentsPackage,
    buildDocumentFragmentPackage,
    buildParagraphOnlyPackage
} from '@ansonlai/docx-redline-js/services/package-builder.js';
import {
    assertRedlineResult,
    INPUT_SANITIZED_WARNING,
    prepareOperationInput
} from './redline-result.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const PART_TYPES = Object.fromEntries(
    ['comments', 'commentsExtended', 'commentsIds', 'commentsExtensible', 'numbering']
        .map(kind => getPartSpec(kind)).map(spec => [`/${spec.path}`, spec.contentType])
);
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

export function isSimplePlainTextRedline(operation) {
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

function packagePart(packageDoc, name) {
    return Array.from(packageDoc.getElementsByTagNameNS(NS_PKG, 'part')).find(
        part => (part.getAttribute('pkg:name') || part.getAttribute('name')) === name
    ) || null;
}

function packagePartXml(packageDoc, name) {
    const xmlData = packagePart(packageDoc, name)?.getElementsByTagNameNS(NS_PKG, 'xmlData')[0];
    const root = Array.from(xmlData?.childNodes || []).find(node => node.nodeType === 1);
    return root ? createSerializer().serializeToString(root) : null;
}

function replacePackagePartXml(packageDoc, name, xml) {
    if (!xml) return;
    let part = packagePart(packageDoc, name);
    if (!part) {
        part = packageDoc.createElementNS(NS_PKG, 'pkg:part');
        part.setAttribute('pkg:name', name);
        part.setAttribute('pkg:contentType', PART_TYPES[name]);
        packageDoc.documentElement.appendChild(part);
    }
    if (PART_TYPES[name]) part.setAttribute('pkg:contentType', PART_TYPES[name]);
    let xmlData = part.getElementsByTagNameNS(NS_PKG, 'xmlData')[0];
    if (!xmlData) {
        xmlData = packageDoc.createElementNS(NS_PKG, 'pkg:xmlData');
        part.appendChild(xmlData);
    }
    while (xmlData.firstChild) xmlData.removeChild(xmlData.firstChild);
    xmlData.appendChild(packageDoc.importNode(parseXmlStrict(xml, name).documentElement, true));
}

function ensurePackageRelationship(packageDoc, kind) {
    const name = '/word/_rels/document.xml.rels';
    let relsXml = packagePartXml(packageDoc, name);
    if (!relsXml) relsXml = `<Relationships xmlns="${NS_REL}"/>`;
    const relsDoc = parseXmlStrict(relsXml, name);
    const relationships = Array.from(relsDoc.getElementsByTagNameNS(NS_REL, 'Relationship'));
    const spec = getPartSpec(kind);
    const type = spec.relType;
    if (!relationships.some(rel => rel.getAttribute('Type') === type)) {
        const used = new Set(relationships.map(rel => rel.getAttribute('Id')));
        let next = 1;
        while (used.has(`rId${next}`)) next++;
        const rel = relsDoc.createElementNS(NS_REL, 'Relationship');
        rel.setAttribute('Id', `rId${next}`);
        rel.setAttribute('Type', type);
        rel.setAttribute('Target', spec.relTarget);
        relsDoc.documentElement.appendChild(rel);
    }
    // Relationship parts have a different content type from the Word XML parts.
    let part = packagePart(packageDoc, name);
    if (!part) {
        part = packageDoc.createElementNS(NS_PKG, 'pkg:part');
        part.setAttribute('pkg:name', name);
        part.setAttribute('pkg:contentType', 'application/vnd.openxmlformats-package.relationships+xml');
        packageDoc.documentElement.appendChild(part);
    }
    let xmlData = part.getElementsByTagNameNS(NS_PKG, 'xmlData')[0];
    if (!xmlData) {
        xmlData = packageDoc.createElementNS(NS_PKG, 'pkg:xmlData');
        part.appendChild(xmlData);
    }
    while (xmlData.firstChild) xmlData.removeChild(xmlData.firstChild);
    xmlData.appendChild(packageDoc.importNode(relsDoc.documentElement, true));
}

function readBatchSource(scopeOoxml) {
    if (scopeOoxml.includes('<pkg:package')) {
        const packageDoc = parseXmlStrict(scopeOoxml, 'Word scope package');
        const documentXml = packagePartXml(packageDoc, '/word/document.xml');
        if (!documentXml) throw new Error('Word scope package has no document.xml part');
        const parts = {
            documentXml,
            commentsXml: packagePartXml(packageDoc, '/word/comments.xml'),
            commentsExtendedXml: packagePartXml(packageDoc, '/word/commentsExtended.xml'),
            commentsIdsXml: packagePartXml(packageDoc, '/word/commentsIds.xml'),
            commentsExtensibleXml: packagePartXml(packageDoc, '/word/commentsExtensible.xml'),
            numberingXml: packagePartXml(packageDoc, '/word/numbering.xml'),
            stylesXml: packagePartXml(packageDoc, '/word/styles.xml')
        };
        const inspection = inspectDocumentParts(parts);
        if (inspection.status === 'error') throw new Error(inspection.error?.message || 'Cannot inspect Word scope');
        return { ...parts, paragraphs: inspection.paragraphs, packageDoc };
    }
    const extracted = extractReplacementNodesFromOoxml(scopeOoxml);
    if (extracted.status === 'error' || !extracted.replacementNodes?.length) {
        throw new Error(extracted.error?.message || 'Word scope has no editable OOXML nodes');
    }
    const documentXml = extracted.sourceType === 'document'
        ? scopeOoxml
        : wrapParagraphNodesAsDocument(extracted.replacementNodes);
    const parts = { documentXml, numberingXml: extracted.numberingXml || null };
    const inspection = inspectDocumentParts(parts);
    if (inspection.status === 'error') throw new Error(inspection.error?.message || 'Cannot inspect Word scope');
    return { ...parts, paragraphs: inspection.paragraphs, packageDoc: null };
}

/**
 * The numbering part of a Word flat-OPC package (e.g. body.getOoxml()), or null.
 * List generation needs the document's real numbering so generated list IDs
 * do not collide with existing ones.
 */
export function readPackageNumberingXml(scopeOoxml) {
    if (typeof scopeOoxml !== 'string' || !scopeOoxml.includes('<pkg:package')) return null;
    return packagePartXml(parseXmlStrict(scopeOoxml, 'Word scope package'), '/word/numbering.xml');
}

/** Capture targeting identities with the same accepted-view inspector used at execution. */
export function captureSourceBaseline(scopeOoxml) {
    return readBatchSource(scopeOoxml).paragraphs.map(paragraph => ({
        index: paragraph.index, exactText: paragraph.exactText,
        fingerprint: paragraph.fingerprint, paragraphId: paragraph.paragraphId,
        inTable: paragraph.inTable
    }));
}

/**
 * Parse one immutable OOXML source, build and execute a canonical operation batch,
 * and prepare the OOXML payload without reading or writing a host document.
 *
 * `operations` may be an array or a function receiving the inspected source.
 * The returned source omits its internal DOM package object and is safe to pass
 * between portable consumers.
 */
export async function prepareCanonicalBatch(scopeOoxml, operations, options = {}) {
    const source = readBatchSource(scopeOoxml);
    const publicSource = {
        documentXml: source.documentXml,
        paragraphs: source.paragraphs,
        commentsXml: source.commentsXml || null,
        numberingXml: source.numberingXml || null,
        stylesXml: source.stylesXml || null
    };
    const batch = typeof operations === 'function'
        ? await operations(publicSource)
        : operations;
    if (!Array.isArray(batch)) throw new TypeError('Batch operations must be an array');
    if (batch.length === 0) {
        return {
            status: 'noop',
            source: publicSource,
            operations: batch,
            result: { status: 'ok', hasChanges: false, results: [], receipts: [] }
        };
    }

    const runner = typeof options.runner === 'function' ? options.runner : applyOperationsToDocumentXml;
    const runtimeContext = {
        commentsXml: source.commentsXml || null,
        commentsExtendedXml: source.commentsExtendedXml || null,
        numberingXml: source.numberingXml || null,
        stylesXml: source.stylesXml || null
    };
    const result = await runner(
        source.documentXml,
        batch,
        options.author || getDefaultAuthor(),
        runtimeContext,
        {
            atomic: true,
            structuredContent: true,
            pairReplacements: true,
            generateRedlines: options.generateRedlines !== false,
            existingRevisions: options.existingRevisions,
            sanitizeInput: options.sanitizeInput === true,
            onInfo: options.onInfo,
            onWarn: options.onWarn
        }
    );

    if (!result || result.status === 'error' || result.status === 'partial' || result.error || result.rolledBack) {
        return { status: 'refused', source: publicSource, operations: batch, result };
    }
    if (!result.hasChanges) {
        return { status: 'noop', source: publicSource, operations: batch, result };
    }
    try {
        return {
            status: 'ready',
            source: publicSource,
            operations: batch,
            result,
            insertionPayload: buildBatchInsertionPackage(source, result)
        };
    } catch (error) {
        return { status: 'error', source: publicSource, operations: batch, result, error };
    }
}

function buildBatchInsertionPackage(source, result) {
    const serializer = createSerializer();
    const outputDoc = parseXmlStrict(result.documentXml, 'batch output');
    normalizeBodySectionOrderStandalone(outputDoc);
    const outputXml = serializer.serializeToString(outputDoc.documentElement);
    const commentSiblings = reconcileCommentSiblingParts({
        commentsXml: result.commentsXml,
        commentsIdsXml: source.commentsIdsXml,
        commentsExtensibleXml: source.commentsExtensibleXml
    });
    let numberingXml = source.numberingXml || null;
    for (const part of result.numberingXmlParts || []) {
        numberingXml = numberingXml ? mergeNumberingXmlBySchemaOrder(numberingXml, part) : part;
    }
    if (source.packageDoc) {
        replacePackagePartXml(source.packageDoc, '/word/document.xml', outputXml);
        if (result.commentsXml) {
            replacePackagePartXml(source.packageDoc, '/word/comments.xml', result.commentsXml);
            ensurePackageRelationship(source.packageDoc, 'comments');
        }
        if (result.commentsExtendedXml) {
            replacePackagePartXml(source.packageDoc, '/word/commentsExtended.xml', result.commentsExtendedXml);
            ensurePackageRelationship(source.packageDoc, 'commentsExtended');
        }
        for (const kind of ['commentsIds', 'commentsExtensible']) {
            const xml = commentSiblings[`${kind}Xml`];
            if (xml) replacePackagePartXml(source.packageDoc, `/${getPartSpec(kind).path}`, xml);
        }
        // A body edit must also repair the commentsExtended type written before 0.8.1.
        const extendedPart = packagePart(source.packageDoc, '/word/commentsExtended.xml');
        if (extendedPart) extendedPart.setAttribute('pkg:contentType', getPartSpec('commentsExtended').contentType);
        if (numberingXml && (result.numberingXmlParts?.length || !source.numberingXml)) {
            replacePackagePartXml(source.packageDoc, '/word/numbering.xml', numberingXml);
            ensurePackageRelationship(source.packageDoc, 'numbering');
        }
        return serializer.serializeToString(source.packageDoc);
    }
    const bodyXml = extractBodyChildElements(outputDoc).map(node => serializer.serializeToString(node)).join('');
    const fragmentPackage = buildDocumentFragmentPackage(bodyXml, {
        includeNumbering: !!numberingXml,
        numberingXml,
        appendTrailingParagraph: false
    });
    if (!result.commentsXml) return fragmentPackage;
    const packageDoc = parseXmlStrict(fragmentPackage, 'batch fragment package');
    replacePackagePartXml(packageDoc, '/word/comments.xml', result.commentsXml);
    ensurePackageRelationship(packageDoc, 'comments');
    if (result.commentsExtendedXml) {
        replacePackagePartXml(packageDoc, '/word/commentsExtended.xml', result.commentsExtendedXml);
        ensurePackageRelationship(packageDoc, 'commentsExtended');
    }
    return serializer.serializeToString(packageDoc);
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
 *   receipt?: Object,
 *   resolvedTarget?: Object,
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
    const prepared = prepareOperationInput(operation, options.sanitizeInput);
    let result = await runner(
        inputDocumentXml,
        prepared.operation,
        options.author || getDefaultAuthor(),
        null,
        {
            generateRedlines: options.generateRedlines !== false,
            sanitizeInput: options.sanitizeInput === true,
            existingRevisions: options.existingRevisions,
            onInfo: options.onInfo,
            onWarn: options.onWarn
        }
    );

    assertRedlineResult(result, 'Shared paragraph operation');
    if (prepared.sanitized) {
        result = {
            ...result,
            warnings: [...(result?.warnings || []), INPUT_SANITIZED_WARNING]
        };
        options.onWarn?.(INPUT_SANITIZED_WARNING);
    }
    if (!result?.hasChanges) {
        return {
            hasChanges: false,
            singleParagraphOutput: false,
            receipt: result?.receipt,
            resolvedTarget: result?.resolvedTarget,
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
        receipt: result.receipt,
        resolvedTarget: result.resolvedTarget,
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
 *   receipt?: Object,
 *   resolvedTarget?: Object,
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
    const prepared = prepareOperationInput(operation, options.sanitizeInput);
    const cleanRangeWorkaround = prepared.operation?.type === 'redline'
        && !!prepared.operation?.targetEndRef
        && options.generateRedlines === false;
    let result = await runner(
        inputDocumentXml,
        prepared.operation,
        options.author || getDefaultAuthor(),
        null,
        {
            generateRedlines: cleanRangeWorkaround ? true : options.generateRedlines !== false,
            sanitizeInput: options.sanitizeInput === true,
            existingRevisions: options.existingRevisions,
            onInfo: options.onInfo,
            onWarn: options.onWarn
        }
    );

    assertRedlineResult(result, 'Shared scope operation');
    if (cleanRangeWorkaround && result?.hasChanges) {
        const accepted = acceptTrackedChangesInOoxml(result.documentXml, {
            author: options.author || getDefaultAuthor()
        });
        assertRedlineResult(accepted, 'Clean range normalization');
        result = {
            ...result,
            documentXml: accepted.oxml,
            warnings: [...(result?.warnings || []), ...(accepted?.warnings || [])]
        };
    }
    if (prepared.sanitized) {
        result = {
            ...result,
            warnings: [...(result?.warnings || []), INPUT_SANITIZED_WARNING]
        };
        options.onWarn?.(INPUT_SANITIZED_WARNING);
    }
    if (!result?.hasChanges) {
        return {
            hasChanges: false,
            singleParagraphOutput: false,
            receipt: result?.receipt,
            resolvedTarget: result?.resolvedTarget,
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
        receipt: result.receipt,
        resolvedTarget: result.resolvedTarget,
        warnings: result.warnings || []
    };
}


export const captureWordSourceBaseline = captureSourceBaseline;

export { planRedlineBatchOperations, assertSourceBaseline } from './redline-plan.js';
export { assertRedlineResult, INPUT_SANITIZED_WARNING, prepareOperationInput, RedlineOperationError } from './redline-result.js';
export { planAgenticListOperations } from '../commands/list-operation-plan.js';
export { validateListRequest } from '../commands/agentic-request-validation.js';
