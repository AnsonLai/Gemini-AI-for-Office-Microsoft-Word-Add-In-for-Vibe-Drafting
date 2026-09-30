import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { configureLogger, openDocx } from '@ansonlai/docx-redline-js';

const BLANK_DOCX = fileURLToPath(new URL('../../assets/blank.docx', import.meta.url));

configureLogger({ log: () => {}, warn: () => {}, error: () => {} });

function assertInspection(doc) {
    const inspection = doc.inspect();
    if (inspection.status === 'error') throw operationError(inspection.error);
    return inspection;
}

function operationError(problem) {
    const error = new Error(problem?.message || 'DOCX operation failed');
    error.code = problem?.code || 'OPERATION_ERROR';
    return error;
}

function assertApplied(result) {
    if (result?.status !== 'error' && !result?.error) return result;
    const problem = result?.results?.find(item => item.error)?.error || result?.error;
    const error = operationError(problem);
    error.details = {
        ...(result?.results ? { results: result.results } : {}),
        ...(result?.receipts ? { receipts: result.receipts } : {}),
        ...(result?.rolledBack !== undefined ? { rolledBack: result.rolledBack } : {})
    };
    throw error;
}

export async function createDocument(title = '') {
    const doc = openDocx(await fs.readFile(BLANK_DOCX));
    if (title) {
        assertApplied(await doc.applyOperations([{
            type: 'replace',
            target: { index: 1, exactText: '' },
            modified: String(title),
            generateRedlines: false
        }], { author: 'MCP AI', atomic: true, structuredContent: false }));
    }
    return doc;
}

export async function openDocument(inputPath) {
    const sourcePath = path.resolve(process.cwd(), inputPath);
    const doc = openDocx(await fs.readFile(sourcePath));
    assertInspection(doc);
    return { doc, sourcePath };
}

export async function saveDocument(doc, outputPath) {
    const resolvedPath = path.resolve(process.cwd(), outputPath);
    const bytes = doc.toUint8Array();
    await fs.writeFile(resolvedPath, bytes);
    return { outputPath: resolvedPath, bytes: bytes.byteLength };
}

export function listDocumentParagraphs(doc, windowing = {}) {
    const paragraphs = assertInspection(doc).paragraphs;
    const start = Math.max(0, Number(windowing.start ?? 0));
    const limit = Math.max(1, Number(windowing.limit ?? 50));
    return {
        total: paragraphs.length,
        start,
        limit,
        items: paragraphs.slice(start, start + limit).map(paragraph => ({
            id: paragraph.paragraphId || `idx:${paragraph.index}`,
            source: paragraph.paragraphId ? 'paraId' : 'index',
            index: paragraph.index,
            text: paragraph.text
        }))
    };
}

export function resolveDocumentParagraph(doc, paragraphId) {
    const paragraphs = assertInspection(doc).paragraphs;
    const byId = paragraphs.find(item => item.paragraphId === paragraphId);
    if (byId) return byId;
    const match = String(paragraphId).match(/^idx:(\d+)$/i);
    if (match) {
        const byIndex = paragraphs[Number(match[1]) - 1];
        if (byIndex) return byIndex;
    }
    throw operationError({ code: 'TARGET_NOT_FOUND', message: `Paragraph not found: ${paragraphId}` });
}

function targetFor(paragraph) {
    return {
        index: paragraph.index,
        exactText: paragraph.exactText,
        ...(paragraph.paragraphId ? { paragraphId: paragraph.paragraphId } : {}),
        ...(paragraph.fingerprint ? { fingerprint: paragraph.fingerprint } : {})
    };
}

export async function applyDocumentOperations(doc, operations, options = {}) {
    return assertApplied(await doc.applyOperations(operations, {
        atomic: true,
        structuredContent: true,
        pairReplacements: true,
        ...options
    }));
}

export async function editDocumentParagraph(doc, paragraphId, newText, options = {}) {
    const paragraph = resolveDocumentParagraph(doc, paragraphId);
    const result = await applyDocumentOperations(doc, [{
        type: 'replace',
        target: targetFor(paragraph),
        modified: String(newText),
        author: options.author || 'MCP AI',
        generateRedlines: options.generateRedlines ?? true
    }]);
    const updated = assertInspection(doc).paragraphs[paragraph.index - 1];
    return {
        result,
        paragraphId: updated?.paragraphId || (updated ? `idx:${updated.index}` : paragraphId),
        updatedText: updated?.text || '',
        changed: Boolean(result.written)
    };
}

export async function addDocumentComment(doc, paragraphId, textToFind, comment, options = {}) {
    const paragraph = resolveDocumentParagraph(doc, paragraphId);
    const beforeCount = assertInspection(doc).comments.length;
    const result = await applyDocumentOperations(doc, [{
        type: 'comment',
        target: targetFor(paragraph),
        textToComment: String(textToFind),
        commentContent: String(comment),
        author: options.author || 'MCP AI'
    }]);
    const commentsApplied = assertInspection(doc).comments.length - beforeCount;
    return { result, paragraphId, commentsApplied };
}
