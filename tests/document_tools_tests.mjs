import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
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
    getGenerateRedlines: () => true,
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

// A source change while checkpoint persistence is in flight refuses before any write.
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

console.log('PASS: direct document tools checkpoint, freshness, single-write, verification, no-op, cancellation and failure receipts');
