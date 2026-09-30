import assert from 'node:assert/strict';
import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import { openDocx } from '@ansonlai/docx-redline-js';

const serverPath = fileURLToPath(new URL('../src/server.mjs', import.meta.url));
const client = new Client({ name: 'wp5-test', version: '1.0.0' }, { capabilities: {} });
const transport = new StdioClientTransport({ command: process.execPath, args: [serverPath], stderr: 'pipe' });
const outputPath = path.join(os.tmpdir(), `mcp-wp5-${crypto.randomUUID()}.docx`);
const newPath = path.join(os.tmpdir(), `mcp-wp5-new-${crypto.randomUUID()}.docx`);

async function call(name, args) {
    const response = await client.callTool({ name, arguments: args });
    return { ...response, body: JSON.parse(response.content[0].text) };
}

try {
    await client.connect(transport);
    const names = (await client.listTools()).tools.map(tool => tool.name);
    assert.ok(names.includes('docx_apply_operations'));
    const immediatelySaved = await call('docx_new', { title: 'Fresh document.', outputPath: newPath });
    assert.equal(immediatelySaved.body.saved.outputPath, newPath);
    const fresh = openDocx(await fs.readFile(newPath));
    assert.equal(fresh.inspect().paragraphs[0].text, 'Fresh document.');
    assert.equal(fresh.inspect().paragraphs[0].hasRevisions, false);
    await call('docx_close', { sessionId: immediatelySaved.body.sessionId });
    const created = await call('docx_new', { title: 'Original clause.' });
    assert.equal(created.body.paragraphs.items[0].text, 'Original clause.');
    const id = created.body.sessionId;
    const handle = created.body.paragraphs.items[0].id;
    const failed = await call('docx_apply_operations', {
        sessionId: id,
        operations: [{ type: 'replace', target: { exactText: 'Missing' }, modified: 'Bad' }]
    });
    assert.equal(failed.isError, true);
    assert.equal(failed.body.code, 'TARGET_NOT_FOUND');
    const afterFailure = await call('docx_list_paragraphs', { sessionId: id });
    assert.equal(afterFailure.body.items[0].text, 'Original clause.');
    const edited = await call('docx_edit_paragraph', { sessionId: id, paragraphId: handle, newText: 'Updated clause.' });
    assert.equal(edited.body.changed, true);
    const batch = await call('docx_apply_operations', {
        sessionId: id,
        operations: [{
            type: 'replace',
            target: { exactText: 'Updated clause.' },
            replacements: [{ find: 'clause', replace: 'agreement' }]
        }]
    });
    assert.equal(batch.body.changed, true);
    assert.equal((await call('docx_list_paragraphs', { sessionId: id })).body.items[0].text, 'Updated agreement.');
    const commented = await call('docx_add_comment', {
        sessionId: id, paragraphId: edited.body.paragraphId, textToFind: 'Updated', comment: 'Review.'
    });
    assert.equal(commented.body.commentsApplied, 1);
    const saved = await call('docx_save_as', { sessionId: id, outputPath });
    assert.equal(saved.body.dirty, false);
    const doc = openDocx(await fs.readFile(outputPath));
    assert.equal(doc.inspect().paragraphs[0].text, 'Updated agreement.');
    assert.equal(doc.inspect().comments.length, 1);
    const reopened = await call('docx_open', { path: outputPath });
    assert.equal(reopened.body.paragraphs.items[0].text, 'Updated agreement.');
    assert.equal((await call('docx_close', { sessionId: id })).body.closed, true);
    assert.equal((await call('docx_close', { sessionId: reopened.body.sessionId })).body.closed, true);
    console.log('PASS: MCP server lifecycle workflow');
} finally {
    await client.close();
    await fs.rm(outputPath, { force: true });
    await fs.rm(newPath, { force: true });
}
