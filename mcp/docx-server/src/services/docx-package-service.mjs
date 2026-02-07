import fs from 'node:fs/promises';
import path from 'node:path';
import JSZip from 'jszip';
import {
    CONTENT_TYPE_COMMENTS,
    CONTENT_TYPE_DOCUMENT,
    CONTENT_TYPE_NUMBERING,
    NS_CT,
    NS_RELS,
    NS_W,
    REL_COMMENTS_TYPE,
    REL_NUMBERING_TYPE,
    REL_OFFICE_DOCUMENT_TYPE,
    allByTag,
    escapeXml,
    firstByTag,
    normalizeBodySectionOrder,
    parseXmlStrict,
    serializeXml
} from './xml-utils.mjs';

const DOCUMENT_PATH = 'word/document.xml';
const CONTENT_TYPES_PATH = '[Content_Types].xml';
const ROOT_RELS_PATH = '_rels/.rels';
const DOC_RELS_PATH = 'word/_rels/document.xml.rels';
const COMMENTS_PATH = 'word/comments.xml';
const NUMBERING_PATH = 'word/numbering.xml';

/**
 * @param {{ title?: string }} [options]
 */
export async function createNewDocxPackage(options = {}) {
    const title = String(options.title || '').trim();
    const zip = new JSZip();

    const documentXml = buildMinimalDocumentXml(title);
    zip.file(CONTENT_TYPES_PATH, buildMinimalContentTypesXml());
    zip.file(ROOT_RELS_PATH, buildRootRelationshipsXml());
    zip.file(DOC_RELS_PATH, buildDocumentRelationshipsXml());
    zip.file(DOCUMENT_PATH, documentXml);

    return {
        zip,
        documentXml: normalizeDocumentXml(documentXml)
    };
}

/**
 * @param {string} inputPath
 */
export async function loadDocxFromPath(inputPath) {
    const resolvedPath = path.resolve(process.cwd(), inputPath);
    const buffer = await fs.readFile(resolvedPath);
    const zip = await JSZip.loadAsync(buffer);
    const documentXml = await readZipText(zip, DOCUMENT_PATH);

    if (!documentXml) {
        throw new Error(`Missing required part: ${DOCUMENT_PATH}`);
    }

    return {
        zip,
        documentXml: normalizeDocumentXml(documentXml),
        sourcePath: resolvedPath
    };
}

/**
 * @param {any} session
 * @param {string} outputPath
 */
export async function saveDocxSessionToPath(session, outputPath) {
    const resolvedPath = path.resolve(process.cwd(), outputPath);
    const normalized = normalizeDocumentXml(session.documentXml);
    session.documentXml = normalized;
    session.zip.file(DOCUMENT_PATH, normalized);

    const data = await session.zip.generateAsync({ type: 'nodebuffer' });
    await fs.writeFile(resolvedPath, data);

    return {
        outputPath: resolvedPath,
        bytes: data.byteLength
    };
}

/**
 * @param {any} zip
 * @param {string|null|undefined} numberingXml
 */
export async function ensureNumberingArtifacts(zip, numberingXml) {
    if (!numberingXml) {
        return { addedNumberingPart: false, hadExistingNumberingPart: false };
    }

    const existing = await readZipText(zip, NUMBERING_PATH);
    let addedPart = false;
    if (!existing) {
        zip.file(NUMBERING_PATH, numberingXml);
        addedPart = true;
    }

    await ensureContentTypeOverride(zip, '/word/numbering.xml', CONTENT_TYPE_NUMBERING);
    await ensureDocumentRelationship(zip, REL_NUMBERING_TYPE, 'numbering.xml');

    return {
        addedNumberingPart: addedPart,
        hadExistingNumberingPart: !!existing
    };
}

/**
 * @param {any} zip
 * @param {string|null|undefined} commentsXml
 */
export async function ensureCommentsArtifacts(zip, commentsXml) {
    if (!commentsXml) {
        return { addedComments: 0 };
    }

    const incomingDoc = parseXmlStrict(commentsXml, COMMENTS_PATH);
    const incomingComments = allByTag(incomingDoc, 'comment');

    const existingText = await readZipText(zip, COMMENTS_PATH);
    if (!existingText) {
        zip.file(COMMENTS_PATH, commentsXml);
        await ensureContentTypeOverride(zip, '/word/comments.xml', CONTENT_TYPE_COMMENTS);
        await ensureDocumentRelationship(zip, REL_COMMENTS_TYPE, 'comments.xml');
        return { addedComments: incomingComments.length };
    }

    const existingDoc = parseXmlStrict(existingText, COMMENTS_PATH);
    const existingRoot = existingDoc.documentElement;
    const existingIds = new Set(
        allByTag(existingDoc, 'comment')
            .map(comment => comment.getAttribute('w:id') || comment.getAttribute('id'))
            .filter(Boolean)
    );

    let added = 0;
    for (const incoming of incomingComments) {
        const id = incoming.getAttribute('w:id') || incoming.getAttribute('id');
        if (id && existingIds.has(id)) {
            throw new Error(`Duplicate comment id detected while merging comments.xml: ${id}`);
        }
        existingRoot.appendChild(existingDoc.importNode(incoming, true));
        if (id) existingIds.add(id);
        added += 1;
    }

    zip.file(COMMENTS_PATH, serializeXml(existingDoc));
    await ensureContentTypeOverride(zip, '/word/comments.xml', CONTENT_TYPE_COMMENTS);
    await ensureDocumentRelationship(zip, REL_COMMENTS_TYPE, 'comments.xml');

    return { addedComments: added };
}

/**
 * @param {string} documentXml
 * @returns {string}
 */
export function normalizeDocumentXml(documentXml) {
    const doc = parseXmlStrict(documentXml, DOCUMENT_PATH);
    normalizeBodySectionOrder(doc);
    return serializeXml(doc);
}

/**
 * @param {any} zip
 * @param {string} partName
 * @param {string} contentType
 */
async function ensureContentTypeOverride(zip, partName, contentType) {
    const existing = await readZipText(zip, CONTENT_TYPES_PATH);
    const contentTypesDoc = existing
        ? parseXmlStrict(existing, CONTENT_TYPES_PATH)
        : parseXmlStrict(buildMinimalContentTypesXml(), CONTENT_TYPES_PATH);
    const root = contentTypesDoc.documentElement;

    const overrides = allByTag(contentTypesDoc, 'Override');
    const hasOverride = overrides.some(override => {
        const current = (override.getAttribute('PartName') || '').toLowerCase();
        return current === partName.toLowerCase();
    });

    if (!hasOverride) {
        const override = contentTypesDoc.createElementNS(NS_CT, 'Override');
        override.setAttribute('PartName', partName);
        override.setAttribute('ContentType', contentType);
        root.appendChild(override);
    }

    zip.file(CONTENT_TYPES_PATH, serializeXml(contentTypesDoc));
}

/**
 * @param {any} zip
 * @param {string} relType
 * @param {string} target
 */
async function ensureDocumentRelationship(zip, relType, target) {
    const existing = await readZipText(zip, DOC_RELS_PATH);
    const relsDoc = existing
        ? parseXmlStrict(existing, DOC_RELS_PATH)
        : parseXmlStrict(buildDocumentRelationshipsXml(), DOC_RELS_PATH);

    const root = firstByTag(relsDoc, 'Relationships') || relsDoc.documentElement;
    const rels = allByTag(relsDoc, 'Relationship');

    const hasRel = rels.some(rel => {
        const type = rel.getAttribute('Type') || '';
        const relTarget = rel.getAttribute('Target') || '';
        return type === relType && relTarget.toLowerCase() === target.toLowerCase();
    });

    if (!hasRel) {
        const nextId = nextRelationshipId(rels);
        const rel = relsDoc.createElementNS(NS_RELS, 'Relationship');
        rel.setAttribute('Id', `rId${nextId}`);
        rel.setAttribute('Type', relType);
        rel.setAttribute('Target', target);
        root.appendChild(rel);
    }

    zip.file(DOC_RELS_PATH, serializeXml(relsDoc));
}

/**
 * @param {Element[]} rels
 * @returns {number}
 */
function nextRelationshipId(rels) {
    let max = 0;
    for (const rel of rels) {
        const raw = rel.getAttribute('Id') || '';
        const match = raw.match(/^rId(\d+)$/i);
        if (match) {
            max = Math.max(max, Number(match[1]));
        }
    }
    return max + 1;
}

/**
 * @param {string} title
 * @returns {string}
 */
function buildMinimalDocumentXml(title) {
    const paragraphXml = title
        ? `<w:p><w:r><w:t xml:space="preserve">${escapeXml(title)}</w:t></w:r></w:p>`
        : '<w:p/>';

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS_W}">
  <w:body>
    ${paragraphXml}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

/**
 * @returns {string}
 */
function buildMinimalContentTypesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${NS_CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${CONTENT_TYPE_DOCUMENT}"/>
</Types>`;
}

/**
 * @returns {string}
 */
function buildRootRelationshipsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_RELS}">
  <Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT_TYPE}" Target="word/document.xml"/>
</Relationships>`;
}

/**
 * @returns {string}
 */
function buildDocumentRelationshipsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${NS_RELS}"></Relationships>`;
}

/**
 * @param {any} zip
 * @param {string} filePath
 * @returns {Promise<string|null>}
 */
async function readZipText(zip, filePath) {
    const file = zip.file(filePath);
    return file ? file.async('string') : null;
}
