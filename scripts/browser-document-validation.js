import { openDocx } from '@ansonlai/docx-redline-js';
import { createDocumentSession } from '../browser-demo/document-session.js';

const assert = (condition, message) => { if (!condition) throw new Error(message); };
const equalBytes = (a, b) => a.length === b.length && a.every((byte, index) => byte === b[index]);
const texts = session => session.inspect().paragraphs.map(p => p.exactText);
const sourcePath = '../tests/fixtures/agentic-lists/nested-lists-source.docx';
const runButton = document.getElementById('run');
const output = document.getElementById('result');

runButton.addEventListener('click', async () => {
    const checks = [];
    runButton.disabled = true;
    try {
        const sourceBytes = new Uint8Array(await (await fetch(sourcePath)).arrayBuffer());
        const source = createDocumentSession(sourceBytes);
        const originalText = texts(source);
        const originalDoc = openDocx(sourceBytes);
        assert(originalText.length === 24, 'Frozen source paragraph count');
        const promptRows = source.getPromptParagraphs();
        assert(/^2\./.test(promptRows.find(p => p.text === 'Number Root B').formattedText.trim()), 'Prompt numbering progression');
        assert(/^\s+1\.1\./.test(promptRows.find(p => p.text === 'Number Nested Anchor').formattedText), 'Prompt nested numbering');
        checks.push('open and inspect Word-authored DOCX');

        const result = await source.applyOperations([{
            type: 'redline', targetRef: 'P2', target: originalText[1],
            replacements: [{ find: 'before', replace: 'beside' }]
        }], { author: 'Browser Integration Test', generateRedlines: true });
        assert(result.status === 'ok' && result.hasChanges, 'Localized edit refused');
        const downloaded = source.toUint8Array();
        const reopened = openDocx(downloaded);
        assert(reopened.inspect().paragraphs[1].exactText === 'Plain paragraph beside bullet list.', 'Localized edit/reopen text');
        for (const part of ['word/styles.xml', 'word/numbering.xml', 'word/_rels/document.xml.rels']) {
            assert(equalBytes(originalDoc.entries.get(part), reopened.entries.get(part)), `Untouched ${part}`);
        }
        checks.push('localized tracked edit, serialize/reopen and untouched parts');
        const accepted = openDocx(downloaded);
        const rejected = openDocx(downloaded);
        assert((await accepted.resolveRevisions('accept', { allAuthors: true })).status === 'ok', 'Accept failure');
        assert((await rejected.resolveRevisions('reject', { allAuthors: true })).status === 'ok', 'Reject failure');
        assert(accepted.inspect().paragraphs[1].exactText === 'Plain paragraph beside bullet list.', 'Accept text');
        assert(JSON.stringify(rejected.inspect().paragraphs.map(p => p.exactText)) === JSON.stringify(originalText), 'Reject source restoration');
        checks.push('exact Accept All and Reject All');

        const beforeRollback = source.toUint8Array();
        const refusal = await source.applyOperations([
            { type: 'replace', target: { index: 2, exactText: 'Plain paragraph beside bullet list.' }, modified: 'Must roll back.' },
            { type: 'replace', target: { index: 99999 }, modified: 'Invalid target.' }
        ], { author: 'Browser Integration Test' });
        assert(refusal.status === 'error' && refusal.rolledBack && !refusal.written, 'Atomic failure contract');
        assert(equalBytes(beforeRollback, source.toUint8Array()), 'Atomic failure changed session bytes');
        checks.push('mixed valid/invalid batch rolls back without mutation');

        const direct = createDocumentSession(sourceBytes);
        const directResult = await direct.applyOperations([{
            type: 'redline', targetRef: 'P2', replacements: [{ find: 'before', replace: 'beside' }]
        }], { author: 'Browser Integration Test', generateRedlines: false });
        assert(directResult.status === 'ok', 'Direct edit failed');
        const directDoc = openDocx(direct.toUint8Array());
        assert(directDoc.inspect().paragraphs[1].exactText === 'Plain paragraph beside bullet list.', 'Direct edit text');
        assert(!directDoc.inspect().paragraphs.some(p => p.hasRevisions), 'Direct edit generated tracked changes');
        checks.push('direct edit produces no new tracked changes');

        const mixedMode = createDocumentSession(downloaded);
        const originalRevisionParagraph = mixedMode.inspect().paragraphs[1];
        assert(originalRevisionParagraph.hasRevisions, 'Tracked fixture has no source revisions');
        const mixedModeResult = await mixedMode.applyOperations([{
            type: 'redline', targetRef: 'P7', replacements: [{ find: 'after', replace: 'beside' }]
        }], { author: 'Direct Writer', generateRedlines: false });
        assert(mixedModeResult.status === 'ok', 'Direct edit with existing revisions failed');
        const mixedModeDoc = openDocx(mixedMode.toUint8Array());
        assert(mixedModeDoc.inspect().paragraphs[1].hasRevisions, 'Direct edit accepted foreign revisions');
        assert(!mixedModeDoc.inspect().paragraphs[6].hasRevisions, 'Direct edit added revisions at its target');
        checks.push('direct mode preserves existing revisions by another author');

        const commentBytes = new Uint8Array(await (await fetch('../tests/fixtures/wp6/threaded-comments.docx')).arrayBuffer());
        const comments = createDocumentSession(commentBytes);
        const beforeComments = comments.inspect().comments;
        const commented = await comments.applyOperations([{
            type: 'comment', targetRef: 'P3', textToComment: 'local law', commentContent: 'Browser review comment.'
        }], { author: 'Browser Integration Test' });
        assert(commented.status === 'ok', 'Comment failed');
        const commentDoc = openDocx(comments.toUint8Array());
        assert(commentDoc.inspect().comments.length === beforeComments.length + 1, 'Comment count/reopen');
        const newComment = commentDoc.inspect().comments.find(c => c.text === 'Browser review comment.');
        assert(newComment?.paragraphIndex === 3 && newComment?.anchoredText === 'local law', 'New comment content or anchor');
        for (const old of beforeComments) {
            const current = commentDoc.inspect().comments.find(c => c.id === old.id);
            assert(current?.text === old.text && current?.parentCommentId === old.parentCommentId && current?.done === old.done, 'Existing comment thread changed');
        }
        checks.push('comment insertion and existing thread preservation after reopen');
        output.textContent = JSON.stringify({ status: 'passed', checks }, null, 2);
    } catch (error) {
        output.textContent = JSON.stringify({ status: 'failed', checks, error: error.message }, null, 2);
    } finally { runButton.disabled = false; }
});
