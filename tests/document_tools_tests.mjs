import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { inspectDocumentParts, acceptTrackedChangesInOoxml } from '@ansonlai/docx-redline-js';
import { buildDocumentFragmentPackage, buildDocumentCommentsPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { inspectWordSource } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { createDocumentInspectionSession } from '../src/taskpane/modules/chat/document-inspection-session.js';
import { createDocumentTools } from '../src/taskpane/modules/chat/document-tools.js';

globalThis.Word = {
  ChangeTrackingMode: { off: 'Off', trackAll: 'TrackAll' },
  InsertLocation: { replace: 'Replace', Replace: 'Replace' }
};

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function escapeXml(value) {
  return String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
}
function packageFor(text, { bold = false } = {}) {
  const runProperties = bold ? '<w:rPr><w:b/></w:rPr>' : '';
  return buildDocumentFragmentPackage(
    `<w:p><w:r>${runProperties}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`,
    { appendTrailingParagraph: false }
  );
}

function packageForParagraphs(texts) {
  const paragraphs = texts.map(text =>
    `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`
  ).join('');
  return buildDocumentFragmentPackage(paragraphs, { appendTrailingParagraph: false });
}

function createHarness(text, options = {}) {
  const counts = { bodyReads: 0, contentWrites: 0, syncCalls: 0, runCalls: 0 };
  const events = [];
  const checkpoints = [];
  const receipts = [];
  const documentId = 'test-document-1';
  let currentXml = options.initialXml || packageFor(text, options.sourceOptions);
  const body = {
    getOoxml() {
      counts.bodyReads++;
      events.push(`getOoxml:${counts.bodyReads}`);
      return { value: currentXml };
    },
    insertOoxml(payload, mode) {
      counts.contentWrites++;
      events.push(`insertOoxml:${mode}`);
      if (options.insertError) throw Object.assign(new Error(options.insertError.message || 'Insert failed.'), options.insertError);
      currentXml = options.transformWrite ? options.transformWrite(payload) : payload;
    }
  };
  let changeTrackingMode = Word.ChangeTrackingMode.trackAll;
  const document = {
    body,
    load(property) { events.push(`load:${property}`); }
  };
  Object.defineProperty(document, 'changeTrackingMode', {
    get() { return changeTrackingMode; },
    set(value) {
      if (value === Word.ChangeTrackingMode.off && options.trackingSetError) throw options.trackingSetError;
      changeTrackingMode = value;
    }
  });
  const context = {
    document,
    async sync() {
      counts.syncCalls++;
      events.push(`sync:${counts.syncCalls}`);
      await options.onSync?.({ currentXml, counts, events });
    }
  };
  const session = createDocumentInspectionSession({
    readSource: async () => {
      const source = inspectWordSource(currentXml);
      if (source.inspection?.status !== 'ok') throw new Error(`Inspection failed: ${JSON.stringify(source.inspection?.error)}`);
      return { documentId, revisionView: 'accepted', parts: source.parts };
    }
  });
  const controller = options.controller || new AbortController();
  const tools = createDocumentTools({
    session,
    getDocumentId: () => options.currentDocumentId?.() ?? documentId,
    getAuthor: () => 'Document Tools Test',
    getGenerateRedlines: () => options.generateRedlines !== false,
    persistCheckpoint: async source => {
      events.push('checkpoint');
      checkpoints.push(source);
      if (options.onCheckpoint) return options.onCheckpoint({ source, currentXml, setCurrent: value => { currentXml = value; }, controller, events });
      if (options.checkpointError) throw options.checkpointError;
      return 1;
    },
    signal: controller.signal,
    run: async callback => {
      counts.runCalls++;
      events.push('run:start');
      const result = await callback(context);
      events.push('run:callback-complete');
      if (options.rejectAfterRun) throw options.rejectAfterRun;
      return result;
    },
    onReceipt: receipt => receipts.push(receipt)
  });

  return {
    tools, counts, events, checkpoints, receipts, controller,
    get currentXml() { return currentXml; },
    set currentXml(value) { currentXml = value; },
    async handleFor(textNeedle) {
      const result = await tools.inspectDocument({ search: textNeedle, limit: 8 });
      assert.equal(result.status, 'ok', JSON.stringify(result));
      const target = result.targets.find(item => item.text.includes(textNeedle));
      assert.ok(target, `inspection returned a handle for ${JSON.stringify(textNeedle)}`);
      return target.handle;
    },
    async handlesFor(textNeedle) {
      const result = await tools.inspectDocument({ search: textNeedle, limit: 8 });
      assert.equal(result.status, 'ok', JSON.stringify(result));
      return result.targets.filter(item => item.text.includes(textNeedle));
    }
  };
}

async function apply(harness, edits) {
  const receipt = await harness.tools.applyDocumentEdits({ edits });
  assert.equal(harness.receipts.at(-1), receipt, 'the same completed receipt is emitted to the caller');
  return receipt;
}

// A real library edit is checkpointed from the exact source snapshot, checked
// fresh before writing, committed once, and verified from a new host read.
{
  const harness = createHarness('Original clause.');
  const target = await harness.handleFor('Original clause.');
  const originalXml = harness.currentXml;
  const receipt = await apply(harness, [{ type: 'replace_text', target, find: 'Original', replace: 'Revised' }]);
  assert.equal(receipt.status, 'ok', JSON.stringify(receipt));
  assert.equal(receipt.success, true);
  assert.equal(receipt.completion, true);
  assert.equal(receipt.verified, true);
  assert.equal(receipt.written, true);
  assert.equal(receipt.writeAttempted, true);
  assert.equal(receipt.checkpointIndex, 1);
  assert.equal(receipt.requestedCount, 1);
  assert.equal(receipt.accountedCount, 1);
  assert.deepEqual(receipt.counters, { sourceReads: 2, contentWrites: 1, verificationReads: 1 });
  assert.equal(harness.counts.contentWrites, 1);
  assert.equal(harness.checkpoints.length, 1);
  assert.equal(harness.checkpoints[0], originalXml, 'checkpoint stores the exact source package read for compilation');
  assert.ok(harness.events.indexOf('checkpoint') < harness.events.findIndex(event => event.startsWith('insertOoxml:')),
    'checkpoint is durable before the one content write');
  const accepted = inspectWordSource(harness.currentXml).inspection;
  assert.equal(accepted.paragraphs[0].exactText, 'Revised clause.');
}

// Character formatting is applied through the library and checked against a
// fresh source read from Word after the commit.
{
  const harness = createHarness('Keep this phrase plain.');
  const target = await harness.handleFor('Keep this phrase plain.');
  const receipt = await apply(harness, [{
    type: 'format_text', target, find: 'phrase', properties: { bold: true }
  }]);
  assert.equal(receipt.success, true, JSON.stringify(receipt));
  assert.equal(receipt.verified, true);
  assert.equal(receipt.written, true);
  assert.deepEqual(receipt.counters, { sourceReads: 2, contentWrites: 1, verificationReads: 1 });
  const parsed = inspectWordSource(harness.currentXml);
  assert.equal(parsed.inspection.paragraphs[0].exactText, 'Keep this phrase plain.');
  assert.match(parsed.parts.documentXml, /<w:b(?:\s[^>]*)?\/?\s*>/,
    'the committed canonical document retains the requested bold property');
}

// Identical paragraphs have no paraId in this source. Targeting the second
// inspected handle proves the verification fallback follows source index when
// matching repeated text, rather than verifying the first lookalike paragraph.
{
  const repeated = 'Same phrase in both paragraphs.';
  const harness = createHarness(repeated, { initialXml: packageForParagraphs([repeated, repeated]) });
  const candidates = await harness.handlesFor('Same phrase');
  assert.equal(candidates.length, 2, 'both identical paragraphs are independently inspectable');
  assert.notEqual(candidates[0].handle, candidates[1].handle);
  const initial = inspectWordSource(harness.currentXml);
  assert.deepEqual(initial.inspection.paragraphs.map(item => item.paragraphId), [null, null],
    'the fixture exercises the canonical index fallback rather than paragraph IDs');

  const receipt = await apply(harness, [{
    type: 'format_text', target: candidates[1].handle, find: 'phrase', properties: { bold: true }
  }]);
  assert.equal(receipt.success, true, JSON.stringify(receipt));
  assert.equal(receipt.verified, true);
  const committed = inspectWordSource(harness.currentXml);
  assert.deepEqual(committed.inspection.paragraphs.map(item => item.exactText), [repeated, repeated]);
  const paragraphs = committed.parts.documentXml.match(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g) || [];
  assert.equal(paragraphs.length, 2, 'both duplicate paragraphs remain in the source');
  assert.doesNotMatch(paragraphs[0], /<w:b(?:\s[^>]*)?\/?\s*>/,
    'the unselected first duplicate remains unformatted');
  assert.match(paragraphs[1], /<w:b(?:\s[^>]*)?\/?\s*>/,
    'the selected second duplicate received formatting');
}

// An external source change during checkpoint persistence refuses before writing.
{
  const harness = createHarness('Original clause.', {
    onCheckpoint({ setCurrent }) {
      setCurrent(packageFor('Concurrent source change.'));
      return 1;
    }
  });
  const target = await harness.handleFor('Original clause.');
  const receipt = await apply(harness, [{ type: 'replace_text', target, find: 'Original', replace: 'Revised' }]);
  assert.equal(receipt.success, false);
  assert.equal(receipt.writeAttempted, false);
  assert.equal(receipt.written, false);
  assert.equal(receipt.mutationOutcome, 'refused');
  assert.equal(receipt.error.code, 'STALE_DOCUMENT_CONTEXT');
  assert.equal(harness.counts.contentWrites, 0);
  assert.equal(harness.events.some(event => event.startsWith('insertOoxml:')), false);
}

// A failed/invalid checkpoint and cancellation during persistence both stop before commit.
{
  const harness = createHarness('Original clause.', { checkpointError: new Error('Disk unavailable.') });
  const target = await harness.handleFor('Original clause.');
  const receipt = await apply(harness, [{ type: 'replace_text', target, find: 'Original', replace: 'Revised' }]);
  assert.equal(receipt.error.code, 'CHECKPOINT_FAILED');
  assert.equal(receipt.writeAttempted, false);
  assert.equal(harness.counts.contentWrites, 0);
}
{
  const controller = new AbortController();
  const harness = createHarness('Original clause.', {
    controller,
    onCheckpoint({ controller: currentController }) { currentController.abort(); return 1; }
  });
  const target = await harness.handleFor('Original clause.');
  const receipt = await apply(harness, [{ type: 'replace_text', target, find: 'Original', replace: 'Revised' }]);
  assert.equal(receipt.error.code, 'REQUEST_CANCELLED');
  assert.equal(receipt.writeAttempted, false);
  assert.equal(harness.counts.contentWrites, 0);
}

// If Word refuses to disable native tracking, the content write never starts;
// the receipt retains the adapter's requirement for explicit inspection.
{
  const harness = createHarness('Original clause.', {
    trackingSetError: Object.assign(new Error('Word tracking mode unavailable.'), { code: 'HOST_TRACKING_MODE_UNAVAILABLE' })
  });
  const target = await harness.handleFor('Original clause.');
  const receipt = await apply(harness, [{ type: 'replace_text', target, find: 'Original', replace: 'Revised' }]);
  assert.equal(receipt.success, false);
  assert.equal(receipt.writeAttempted, false);
  assert.equal(receipt.written, false);
  assert.equal(receipt.requiresInspection, true);
  assert.equal(harness.counts.contentWrites, 0);
}
{
  const controller = new AbortController();
  controller.abort();
  const harness = createHarness('Original clause.', { controller });
  const receipt = await apply(harness, [{ type: 'rewrite_paragraph', target: 'expired', text: 'Changed.' }]);
  assert.equal(receipt.error.code, 'REQUEST_CANCELLED');
  assert.equal(harness.counts.runCalls, 0);
  assert.equal(harness.counts.contentWrites, 0);
}

// A committed but altered Word result is reported as uncertain and is not replayed.
{
  const harness = createHarness('Original clause.', {
    transformWrite: () => packageFor('Word changed the committed result.')
  });
  const target = await harness.handleFor('Original clause.');
  const receipt = await apply(harness, [{ type: 'replace_text', target, find: 'Original', replace: 'Revised' }]);
  assert.equal(receipt.success, false);
  assert.equal(receipt.status, 'error');
  assert.equal(receipt.writeAttempted, true);
  assert.equal(receipt.written, true);
  assert.equal(receipt.verified, false);
  assert.equal(receipt.mutationOutcome, 'applied_with_host_error');
  assert.equal(receipt.error.code, 'HOST_VERIFICATION_FAILED');
  assert.equal(harness.counts.contentWrites, 1);
}

// Exact, already-satisfied desired state is verified as a no-op with no checkpoint or content write.
{
  const harness = createHarness('Already bold.', { sourceOptions: { bold: true } });
  const target = await harness.handleFor('Already bold.');
  const receipt = await apply(harness, [{ type: 'format_text', target, find: 'Already bold.', properties: { bold: true } }]);
  assert.equal(receipt.success, true);
  assert.equal(receipt.verified, true);
  assert.equal(receipt.written, false);
  assert.equal(receipt.writeAttempted, false);
  assert.equal(receipt.mutationOutcome, 'noop');
  assert.equal(receipt.checkpointIndex, undefined);
  assert.deepEqual(receipt.counters, { sourceReads: 1, contentWrites: 0, verificationReads: 0 });
  assert.equal(harness.checkpoints.length, 0);
}

// One malformed intent rejects the entire requested mixture before checkpoints or writes.
{
  const harness = createHarness('Original clause.');
  const target = await harness.handleFor('Original clause.');
  const receipt = await apply(harness, [
    { type: 'replace_text', target, find: 'Original', replace: 'Revised' },
    { type: 'format_text', target, find: 'Original', properties: { color: 'red' } }
  ]);
  assert.equal(receipt.success, false);
  assert.equal(receipt.writeAttempted, false);
  assert.equal(receipt.written, false);
  assert.equal(receipt.mutationOutcome, 'refused');
  assert.equal(harness.checkpoints.length, 0);
  assert.equal(harness.counts.contentWrites, 0);
}

// A Word.run wrapper that rejects after its callback has verified a commit must
// still preserve the already-observed write receipt for the caller.
{
  const harness = createHarness('Original clause.', { rejectAfterRun: new Error('Word.run post-commit rejection.') });
  const target = await harness.handleFor('Original clause.');
  const receipt = await apply(harness, [{ type: 'replace_text', target, find: 'Original', replace: 'Revised' }]);
  assert.equal(receipt.written, true);
  assert.equal(receipt.writeAttempted, true);
  assert.equal(receipt.mutationOutcome, 'applied_with_host_error');
  assert.equal(receipt.error.code, 'WORD_RUN_POST_COMMIT_ERROR');
  assert.equal(receipt.verified, false);
  assert.equal(harness.counts.contentWrites, 1);
  assert.equal(inspectWordSource(harness.currentXml).inspection.paragraphs[0].exactText, 'Revised clause.');
}

// Existing comments, replies and resolved state survive a combined rewrite
// and new comment, followed by another addition against freshly inspected text.
for (const metadata of [
  '<w15:commentEx w15:paraId="2222BBBB" w15:paraIdParent="1111AAAA"/>',
  '<w15:commentEx w15:paraId="2222BBBB" w15:done="1"/>',
  '<w15:commentEx w15:paraId="2222BBBB" w15:done="0"/>',
]) {
  const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const w15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
  let threaded = buildDocumentCommentsPackage(`<w:document xmlns:w="${W}"><w:body><w:p><w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/><w:r><w:t>Original clause.</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p></w:body></w:document>`,
    `<w:comments xmlns:w="${W}" xmlns:w14="${w14}"><w:comment w:id="0" w:author="Reviewer"><w:p w14:paraId="1111AAAA"><w:r><w:t>Original comment.</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="Responder"><w:p w14:paraId="2222BBBB"><w:r><w:t>Reply or resolved comment.</w:t></w:r></w:p></w:comment></w:comments>`);
  threaded = threaded.replace('</pkg:package>', `<pkg:part pkg:name="/word/commentsExtended.xml" pkg:contentType="application/vnd.ms-word.commentsExtended+xml"><pkg:xmlData><w15:commentsEx xmlns:w15="${w15}"><w15:commentEx w15:paraId="1111AAAA"/>${metadata}</w15:commentsEx></pkg:xmlData></pkg:part></pkg:package>`);
  const inspection = inspectWordSource(threaded).inspection;
  assert.equal(inspection.comments.length, 2);
  const harness = createHarness('Original clause.', { initialXml: threaded });
  const target = await harness.handleFor('Original clause.');
  const receipt = await apply(harness, [
    { type: 'rewrite_paragraph', target, text: 'Revised clause.' },
    { type: 'comment', target, anchor: 'clause.', text: 'New comment.' },
  ]);
  assert.equal(receipt.success, true, JSON.stringify(receipt));
  assert.equal(receipt.verified, true);
  assert.equal(receipt.requestedCount, 2);
  assert.equal(receipt.accountedCount, 2);
  assert.equal(receipt.writeAttempted, true);
  assert.equal(harness.counts.contentWrites, 1);
  assert.equal(harness.checkpoints.length, 1);
  const committed = inspectWordSource(harness.currentXml).inspection;
  assert.equal(committed.paragraphs[0].exactText, 'Revised clause.');
  assert.equal(committed.comments.length, 3);
  const identity = comment => [comment.id, comment.author, comment.text, comment.done, comment.parentCommentId];
  for (const comment of inspection.comments) {
    const retained = committed.comments.find(item => item.id === comment.id);
    assert.deepEqual(identity(retained), identity(comment));
    assert.equal(retained.anchoredText, 'Revised clause.', 'original comment ranges still surround the revised wording');
  }
  const freshTarget = await harness.handleFor('Revised clause.');
  const second = await apply(harness, [{ type: 'comment', target: freshTarget, anchor: 'Revised', text: 'Another comment.' }]);
  assert.equal(second.success, true, JSON.stringify(second));
  assert.equal(second.verified, true);
  assert.equal(harness.counts.contentWrites, 2);
  assert.equal(harness.checkpoints.length, 2);
  const finalComments = inspectWordSource(harness.currentXml).inspection.comments;
  assert.equal(finalComments.length, 4, 'later comment additions retain all previous comments');
  for (const comment of committed.comments) assert.deepEqual(identity(finalComments.find(item => item.id === comment.id)), identity(comment));
  if (metadata.includes('paraIdParent') || metadata.includes('done="1"')) {
    const damaged = createHarness('Original clause.', { initialXml: threaded,
      transformWrite: xml => metadata.includes('paraIdParent')
        ? xml.replace(' w15:paraIdParent="1111AAAA"', '')
        : xml.replace('w15:done="1"', 'w15:done="0"') });
    const damagedTarget = await damaged.handleFor('Original clause.');
    const unverified = await apply(damaged, [{ type: 'comment', target: damagedTarget, anchor: 'clause.', text: 'New comment.' }]);
    assert.equal(unverified.success, false, 'lost reply/resolved state cannot be reported as a successful comment addition');
    assert.equal(unverified.verified, false);
    assert.equal(unverified.written, true);
    assert.equal(unverified.requiresInspection, true);
    assert.equal(damaged.counts.contentWrites, 1, 'an unverified host write must not be replayed');
  }
}

// A normal-run literal edit preserves a foreign pending insertion. An adjacent
// clean paragraph may be edited in the same atomic batch.
{
  const revised = buildDocumentFragmentPackage(`<w:p><w:r><w:t>Base before </w:t></w:r><w:ins w:id="21" w:author="Prior Reviewer" w:date="2026-10-02T00:00:00Z"><w:r><w:t>Pending clause.</w:t></w:r></w:ins><w:r><w:t> base after.</w:t></w:r></w:p><w:p><w:r><w:t>Untouched clause.</w:t></w:r></w:p>`, { appendTrailingParagraph: false });
  const harness = createHarness('Pending clause.', { initialXml: revised, generateRedlines: false });
  const target = await harness.handleFor('Pending clause.');
  const other = await harness.handleFor('Untouched clause.');
  const receipt = await apply(harness, [
    { type: 'replace_text', target, find: 'Base before', replace: 'Revised before' },
    { type: 'replace_text', target: other, find: 'Untouched', replace: 'Changed' },
  ]);
  assert.equal(receipt.success, true, JSON.stringify(receipt));
  assert.equal(receipt.verified, true);
  assert.equal(receipt.written, true);
  assert.equal(receipt.accountedCount, 2);
  assert.equal(receipt.writeAttempted, true);
  assert.equal(receipt.requestedCount, 2);
  assert.equal(harness.counts.contentWrites, 1, 'the foreign-revision target and clean paragraph commit atomically in one write');
  assert.equal(harness.checkpoints.length, 1);
  const beforeParts = inspectWordSource(revised).parts;
  const after = inspectWordSource(harness.currentXml);
  assert.deepEqual(after.inspection.paragraphs.map(paragraph => paragraph.exactText), [
    'Revised before Pending clause. base after.', 'Changed clause.'
  ]);
  const originalInsertion = beforeParts.documentXml.match(/<w:ins\b[\s\S]*?<\/w:ins>/)?.[0];
  const foreignInsertion = after.parts.documentXml.match(/<w:ins\b[\s\S]*?<\/w:ins>/)?.[0];
  assert.ok(foreignInsertion, 'the pending foreign insertion remains tracked');
  assert.equal(foreignInsertion, originalInsertion,
    'an adjacent normal-run edit leaves the foreign reviewer insertion XML exactly unchanged');
  assert.deepEqual(after.inspection.paragraphs[0].revisionAuthors, ['Prior Reviewer']);
  const beforeRejected = inspectDocumentParts(beforeParts, { revisionView: 'rejected' });
  const afterRejected = inspectDocumentParts(after.parts, { revisionView: 'rejected' });
  assert.equal(afterRejected.paragraphs[0].exactText,
    beforeRejected.paragraphs[0].exactText.replace('Base before', 'Revised before'),
    'rejecting the foreign insertion keeps only the requested normal-run change in the base text');
}

// Host verification must detect a flattened review carrier even when accepted
// text is identical, while tolerating Word's semantic-preserving ID/run changes.
{
  const source = buildDocumentFragmentPackage(`<w:p><w:r><w:t>Base before </w:t></w:r><w:ins w:id="21" w:author="Prior Reviewer"><w:r><w:t>Pending clause.</w:t></w:r></w:ins></w:p>`, { appendTrailingParagraph: false });
  const transformations = [
    { label: 'flattened', verified: false,
      transform: xml => xml.replace(/<w:ins\b[^>]*>([\s\S]*?)<\/w:ins>/, '$1') },
    { label: 'renumbered', verified: true,
      transform: xml => xml.replace('w:id="21"', 'w:id="31"') },
    { label: 'split', verified: true,
      transform: xml => xml.replace(/<w:ins\b[^>]*><w:r><w:t>Pending clause\.<\/w:t><\/w:r><\/w:ins>/,
        '<w:ins w:id="31" w:author="Prior Reviewer"><w:r><w:t>Pending </w:t></w:r></w:ins><w:ins w:id="32" w:author="Prior Reviewer"><w:r><w:t>clause.</w:t></w:r></w:ins>') },
    { label: 'reattributed', verified: false,
      transform: xml => xml.replace('w:author="Prior Reviewer"', 'w:author="Different Reviewer"') },
  ];
  for (const { label, verified, transform } of transformations) {
    const harness = createHarness('Pending clause.', { initialXml: source,
      generateRedlines: false, transformWrite: transform });
    const target = await harness.handleFor('Pending clause.');
    const receipt = await apply(harness, [{ type: 'replace_text', target, find: 'Base before', replace: 'Revised before' }]);
    assert.equal(receipt.success, verified, `${label}: ${JSON.stringify(receipt)}`);
    assert.equal(receipt.written, true);
    assert.equal(receipt.verified, verified);
    assert.equal(harness.counts.contentWrites, 1);
    if (!verified) assert.equal(receipt.requiresInspection, true);
    assert.equal(inspectWordSource(harness.currentXml).inspection.paragraphs[0].exactText,
      'Revised before Pending clause.', `${label}: accepted text alone cannot distinguish carrier loss`);
  }
}

// Published 0.8.4 incorrectly tracked a pure insertion in surgical mode even with
// tracking disabled, so the add-in rejected that prepared output (TRACKING_POLICY_MISMATCH).
// 0.9.0 fixed this: the untracked insertion is written as plain text and the foreign
// pending insertion is left untouched. The add-in's guard remains as a defensive check.
{
  const source = buildDocumentFragmentPackage(`<w:p><w:r><w:t>This agreement is governed by local law.</w:t></w:r><w:ins w:id="21" w:author="Prior Reviewer"><w:r><w:t xml:space="preserve"> Pending clause.</w:t></w:r></w:ins></w:p>`, { appendTrailingParagraph: false });
  const harness = createHarness('governed by local law', { initialXml: source, generateRedlines: false });
  const target = await harness.handleFor('governed by local law');
  const receipt = await apply(harness, [{ type: 'replace_text', target,
    find: 'governed by local law', replace: 'governed by applicable local law' }]);
  assert.equal(receipt.success, true, JSON.stringify(receipt));
  assert.equal(receipt.written, true);
  assert.equal(receipt.verified, true);
  assert.equal(harness.counts.contentWrites, 1);
  assert.equal(inspectWordSource(harness.currentXml).inspection.paragraphs[0].exactText,
    'This agreement is governed by applicable local law. Pending clause.');
  assert.equal((harness.currentXml.match(/<w:ins /g) || []).length, 1, 'only the prior reviewer insertion may remain tracked');
}

// A receipt may allocate IDs while the library slices a foreign insertion.
// Those IDs can identify cloned foreign carriers rather than new user-tracked
// text, so an untracked replacement inside the insertion must be judged by its
// resulting carrier/author semantics instead of the receipt ID alone.
{
  const source = buildDocumentFragmentPackage(`<w:p><w:ins w:id="21" w:author="Prior Reviewer" w:date="2026-10-02T00:00:00Z"><w:r><w:t>Pending clause.</w:t></w:r></w:ins></w:p>`, { appendTrailingParagraph: false });
  const harness = createHarness('Pending clause.', { initialXml: source, generateRedlines: false });
  const target = await harness.handleFor('Pending clause.');
  const receipt = await apply(harness, [{ type: 'replace_text', target, find: 'Pending', replace: 'Reviewed' }]);
  assert.equal(receipt.success, true, JSON.stringify(receipt));
  assert.equal(receipt.verified, true);
  assert.equal(receipt.written, true);
  assert.equal(harness.counts.contentWrites, 1);

  const committed = inspectWordSource(harness.currentXml);
  assert.equal(committed.inspection.paragraphs[0].exactText, 'Reviewed clause.');
  const insertion = committed.parts.documentXml.match(/<w:ins\b[\s\S]*?<\/w:ins>/)?.[0];
  assert.ok(insertion, 'the unaffected suffix remains a pending insertion');
  assert.match(insertion, /w:author="Prior Reviewer"/);
  assert.match(insertion, /w:date="2026-10-02T00:00:00Z"/);
  assert.match(insertion, /<w:t(?:\s[^>]*)?> clause\.<\/w:t>/,
    'the unaffected suffix stays inside the original reviewer carrier');
  assert.doesNotMatch(insertion, /Reviewed/,
    'the requested replacement is ordinary untracked text');
  assert.match(committed.parts.documentXml, /<w:r><w:t>Reviewed<\/w:t><\/w:r>/);
}

// A normal run can be formatted while a neighboring foreign insertion remains
// untouched, with redlines disabled and the prior reviewer carrier preserved.
{
  const source = buildDocumentFragmentPackage(`<w:p><w:r><w:t xml:space="preserve">Keep this phrase plain. </w:t></w:r><w:ins w:id="22" w:author="Prior Reviewer" w:date="2026-10-02T00:00:00Z"><w:r><w:t>Pending clause.</w:t></w:r></w:ins></w:p>`, { appendTrailingParagraph: false });
  const harness = createHarness('Keep this phrase plain.', { initialXml: source, generateRedlines: false });
  const target = await harness.handleFor('Keep this phrase plain.');
  const receipt = await apply(harness, [{ type: 'format_text', target, find: 'phrase', properties: { bold: true } }]);
  assert.equal(receipt.success, true, JSON.stringify(receipt));
  assert.equal(receipt.verified, true);
  assert.equal(harness.counts.contentWrites, 1);
  const committed = inspectWordSource(harness.currentXml);
  assert.equal(committed.inspection.paragraphs[0].exactText, 'Keep this phrase plain. Pending clause.');
  assert.deepEqual(committed.inspection.paragraphs[0].revisionAuthors, ['Prior Reviewer']);
  const insertion = committed.parts.documentXml.match(/<w:ins\b[\s\S]*?<\/w:ins>/)?.[0];
  assert.ok(insertion);
  assert.match(insertion, /w:id="22"/);
  assert.match(insertion, /w:author="Prior Reviewer"/);
  assert.doesNotMatch(insertion, /<w:b(?:\s[^>]*)?\/?\s*>/,
    'formatting the normal run does not alter the neighboring tracked insertion');
  const boldRun = (committed.parts.documentXml.match(/<w:r\b[\s\S]*?<\/w:r>/g) || [])
    .find(run => run.includes('phrase') && /<w:b(?:\s[^>]*)?\/?\s*>/.test(run));
  assert.ok(boldRun, 'the requested normal run receives bold formatting');
}

// Revision policy must not force surgical text mode for an ordinary structural
// list operation; its complete-list recipe must remain a real two-item list.
{
  const numberingXml = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
  const listSource = buildDocumentFragmentPackage('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Existing list item.</w:t></w:r></w:p>',
    { appendTrailingParagraph: false, includeNumbering: true, numberingXml });
  const harness = createHarness('Existing list item.', { initialXml: listSource });
  const target = await harness.handleFor('Existing list item.');
  const receipt = await apply(harness, [{ type: 'insert_list_item', target, text: 'Added list item.' }]);
  assert.equal(receipt.success, true, JSON.stringify(receipt));
  assert.equal(receipt.verified, true);
  const parts = inspectWordSource(harness.currentXml).parts;
  const accepted = acceptTrackedChangesInOoxml(parts.documentXml, { allAuthors: true });
  assert.ok(accepted.acceptedCount > 0, JSON.stringify(accepted));
  const committed = inspectDocumentParts({ ...parts, documentXml: accepted.oxml }).paragraphs;
  assert.deepEqual(committed.map(paragraph => paragraph.exactText), ['Existing list item.', 'Added list item.']);
  assert.ok(committed.every(paragraph => paragraph.list?.format === 'bullet'));
  assert.equal(harness.counts.contentWrites, 1);
}

// Own-only revisions retain the documented same-author baseline merge policy;
// they must not be forced into cross-author surgical mode by their presence.
{
  const source = buildDocumentFragmentPackage(`<w:p><w:r><w:t>This agreement is governed by local law.</w:t></w:r><w:ins w:id="21" w:author="Document Tools Test"><w:r><w:t xml:space="preserve"> Prior own change.</w:t></w:r></w:ins></w:p>`, { appendTrailingParagraph: false });
  const harness = createHarness('governed by local law', { initialXml: source, generateRedlines: false });
  const target = await harness.handleFor('governed by local law');
  const receipt = await apply(harness, [{ type: 'replace_text', target,
    find: 'governed by local law', replace: 'governed by applicable local law' }]);
  assert.equal(receipt.success, true, JSON.stringify(receipt));
  assert.equal(receipt.written, true);
  const actual = inspectWordSource(harness.currentXml);
  assert.equal(actual.inspection.paragraphs[0].exactText,
    'This agreement is governed by applicable local law. Prior own change.');
  assert.doesNotMatch(actual.parts.documentXml, /<w:(?:ins|del)\b/,
    'the same-author baseline merge produces clean text with tracking disabled');
}

// The public default merge-same-author policy intentionally merges the user's
// own pending edit with its baseline when the request explicitly asks for an
// untracked whole-paragraph rewrite.
{
  const ownRevision = buildDocumentFragmentPackage(`<w:p><w:r><w:t>Base wording </w:t></w:r><w:ins w:id="23" w:author="Document Tools Test" w:date="2026-10-02T00:00:00Z"><w:r><w:t>own pending change.</w:t></w:r></w:ins></w:p>`, { appendTrailingParagraph: false });
  const harness = createHarness('Base wording own pending change.', { initialXml: ownRevision, generateRedlines: false });
  const target = await harness.handleFor('Base wording own pending change.');
  const receipt = await apply(harness, [{ type: 'rewrite_paragraph', target, text: 'Clean final wording.' }]);
  assert.equal(receipt.success, true, JSON.stringify(receipt));
  assert.equal(receipt.verified, true);
  assert.equal(harness.counts.contentWrites, 1);
  const committed = inspectWordSource(harness.currentXml);
  assert.equal(committed.inspection.paragraphs[0].exactText, 'Clean final wording.');
  assert.deepEqual(committed.inspection.paragraphs[0].revisionAuthors, [],
    'the intentional same-author clean rewrite leaves no tracked revision');
  assert.doesNotMatch(committed.parts.documentXml, /<w:(?:ins|del)\b/);
}

console.log('PASS: direct document tools checkpoint, freshness, single-write, verification, no-op, cancellation and failure receipts');
