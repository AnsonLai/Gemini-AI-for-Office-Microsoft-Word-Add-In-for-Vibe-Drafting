import './setup-xml-provider.mjs';

import assert from 'assert';
import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { validateDocxPackage } from '@ansonlai/docx-redline-js';

import {
    createNewDocxPackage,
    loadDocxFromPath,
    saveDocxSessionToPath
} from '../mcp/docx-server/src/services/docx-package-service.mjs';
import { DocxSessionStore } from '../mcp/docx-server/src/services/docx-session-store.mjs';
import {
    listParagraphs,
    replaceParagraph,
    resolveParagraph,
    serializeParagraph
} from '../mcp/docx-server/src/services/paragraph-targeting-service.mjs';
import {
    deriveParagraphAcceptedText,
    reconcileParagraphEdit
} from '../mcp/docx-server/src/services/docx-redline-js-service.mjs';

async function run() {
    const outputPath = path.join(os.tmpdir(), `docx-redline-mcp-smoke-${crypto.randomUUID()}.docx`);
    try {
        const created = await createNewDocxPackage({ title: 'Original clause.' });
        const sessions = new DocxSessionStore();
        const session = sessions.create({
            ...created,
            defaultGenerateRedlines: true
        });

        const beforeList = listParagraphs(session.documentXml);
        assert.strictEqual(beforeList.total, 1);
        assert.strictEqual(beforeList.items[0].text, 'Original clause.');

        const beforeFailureXml = session.documentXml;
        await assert.rejects(
            reconcileParagraphEdit({
                paragraphXml: serializeParagraph(resolveParagraph(session.documentXml, 'idx:1').paragraph),
                paragraphText: 'Stale target text.',
                modifiedText: 'This must not commit.',
                author: 'MCP Smoke',
                generateRedlines: true
            }),
            error => error.code === 'TARGET_NOT_FOUND'
        );
        assert.strictEqual(session.documentXml, beforeFailureXml);
        assert.strictEqual(session.dirty, false);

        const resolved = resolveParagraph(session.documentXml, 'idx:1');
        const paragraphXml = serializeParagraph(resolved.paragraph);
        const reconciliation = await reconcileParagraphEdit({
            paragraphXml,
            paragraphText: deriveParagraphAcceptedText(paragraphXml),
            modifiedText: 'Updated clause.',
            author: 'MCP Smoke',
            generateRedlines: true
        });
        assert.strictEqual(reconciliation.hasChanges, true);

        session.documentXml = replaceParagraph(
            resolved.doc,
            resolved.paragraph,
            reconciliation.replacementNodes
        );
        session.dirty = true;
        sessions.touch(session);

        const saved = await saveDocxSessionToPath(session, outputPath);
        assert.ok(saved.bytes > 0);
        session.dirty = false;

        const reopened = await loadDocxFromPath(outputPath);
        await validateDocxPackage(reopened.zip);
        const reopenedParagraph = resolveParagraph(reopened.documentXml, 'idx:1').paragraph;
        assert.strictEqual(
            deriveParagraphAcceptedText(serializeParagraph(reopenedParagraph)),
            'Updated clause.'
        );
        assert.ok(reopened.documentXml.includes('<w:ins'));
        assert.ok(reopened.documentXml.includes('<w:del'));

        assert.strictEqual(sessions.close(session.sessionId), true);
        console.log('PASS: MCP DOCX create/edit/save/reopen smoke test');
    } finally {
        await fs.rm(outputPath, { force: true });
    }
}

run().catch(error => {
    console.error('FAIL:', error?.stack || error);
    process.exit(1);
});
