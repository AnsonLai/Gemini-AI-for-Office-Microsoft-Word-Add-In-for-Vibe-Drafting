import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { validateDocxPackage } from '@ansonlai/docx-redline-js';
import { MemoryZip, unzipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { DocxSessionStore } from '../mcp/docx-server/src/services/docx-session-store.mjs';
import {
    addDocumentComment,
    applyDocumentOperations,
    createDocument,
    editDocumentParagraph,
    listDocumentParagraphs,
    openDocument,
    saveDocument
} from '../mcp/docx-server/src/services/docx-document-service.mjs';

async function run() {
    const outputPath = path.join(os.tmpdir(), `docx-redline-mcp-smoke-${crypto.randomUUID()}.docx`);
    try {
        const doc = await createDocument('Original clause.');
        const sessions = new DocxSessionStore();
        const session = sessions.create({ doc, defaultGenerateRedlines: true });
        const before = listDocumentParagraphs(doc);
        assert.equal(before.total, 1);
        assert.equal(before.items[0].text, 'Original clause.');
        assert.equal(doc.inspect().paragraphs[0].hasRevisions, false);

        const beforeBytes = doc.toUint8Array();
        await assert.rejects(
            applyDocumentOperations(doc, [{ type: 'replace', target: { exactText: 'Stale target text.' }, modified: 'This must not commit.' }]),
            error => error.code === 'TARGET_NOT_FOUND'
        );
        assert.deepEqual(doc.toUint8Array(), beforeBytes);
        assert.equal(session.dirty, false);

        const edit = await editDocumentParagraph(doc, before.items[0].id, 'Updated clause.', { author: 'MCP Smoke' });
        assert.equal(edit.changed, true);
        assert.equal(edit.updatedText, 'Updated clause.');
        session.dirty = true;
        sessions.touch(session);
        const comment = await addDocumentComment(doc, edit.paragraphId, 'Updated', 'Review this word.', { author: 'Reviewer' });
        assert.equal(comment.commentsApplied, 1);

        const saved = await saveDocument(doc, outputPath);
        assert.ok(saved.bytes > 0);
        session.dirty = false;
        const reopened = await openDocument(outputPath);
        assert.equal(reopened.doc.inspect().paragraphs[0].text, 'Updated clause.');
        assert.equal(reopened.doc.inspect().comments.length, 1);
        assert.equal(reopened.doc.inspect().paragraphs[0].hasRevisions, true);
        await validateDocxPackage(new MemoryZip(unzipDocx(reopened.doc.toUint8Array())));
        assert.equal(sessions.close(session.sessionId), true);
        console.log('PASS: MCP DOCX create/edit/comment/save/reopen smoke test');
    } finally {
        await fs.rm(outputPath, { force: true });
    }
}

run().catch(error => {
    console.error('FAIL:', error?.stack || error);
    process.exit(1);
});
