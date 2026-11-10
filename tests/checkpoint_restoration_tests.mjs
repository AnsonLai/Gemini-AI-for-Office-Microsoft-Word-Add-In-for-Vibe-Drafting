import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import { buildDocumentCommentsPackage, buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { inspectWordSource } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { computeWordRestorationToken, computeWordRevertGuardToken } from '../src/taskpane/modules/docx-redline-js-integration/word-source-identity.js';
import { restoreBodyCheckpoint, CHANGED_SINCE_EDIT_MESSAGE } from '../src/taskpane/modules/chat/checkpoint-restoration.js';

globalThis.Word = { ChangeTrackingMode: { off: 'Off', trackAll: 'TrackAll' } };
const snapshot = buildDocumentFragmentPackage('<w:p><w:r><w:t>Saved provision.</w:t></w:r></w:p>', { appendTrailingParagraph: false });
function harness({ checkpoint = snapshot, failureSync, transform, postRunError, documentChanged, controller, cancelDuringCommit, currentXml, guardToken } = {}) {
  let current = currentXml ?? checkpoint, syncs = 0, writes = 0, identity = 'document';
  const modes = [];
  let mode = 'TrackAll';
  const document = { load() {}, body: {
    getOoxml: () => ({ value: current }),
    insertOoxml(value, location) { assert.equal(location, 'Replace'); writes++; current = transform ? transform(value) : value; },
    clear() { assert.fail('restoration must replace once without clearing'); },
  } };
  Object.defineProperty(document, 'changeTrackingMode', { get: () => mode, set(value) { modes.push(value); mode = value; } });
  const context = { document, async sync() {
    syncs++;
    if (cancelDuringCommit && syncs === 3) controller.abort();
    if (documentChanged && syncs === 2) identity = 'different-document';
    if (syncs === failureSync) throw new Error('Host synchronization failed.');
  } };
  return {
    async restore(value = checkpoint) { return restoreBodyCheckpoint(value, {
      getDocumentId: () => identity, signal: controller?.signal,
      ...(guardToken !== undefined ? { guardToken } : {}),
      run: async callback => { const result = await callback(context); if (postRunError) throw new Error('Word.run failed after callback.'); return result; },
    }); },
    state: () => ({ writes, mode, modes }),
  };
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
function addFlatOpcPart(flatOpc, name, xml) {
  const anchor = '<pkg:part pkg:name="/word/document.xml"';
  assert.ok(flatOpc.includes(anchor));
  const part = `<pkg:part pkg:name="/word/${name}" pkg:contentType="application/xml"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`;
  return flatOpc.replace(anchor, `${part}\n  ${anchor}`);
}
function buildThreadedCheckpoint() {
  const documentXml = `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>
    <w:p w14:paraId="10000001"><w:commentRangeStart w:id="1"/><w:r><w:t>Root anchor text.</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>
    <w:p w14:paraId="10000002"><w:commentRangeStart w:id="2"/><w:r><w:t>Reply anchor text.</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>
    <w:p w14:paraId="10000003"><w:commentRangeStart w:id="3"/><w:r><w:t>Other root anchor.</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r></w:p>
  </w:body></w:document>`;
  const commentsXml = `<w:comments xmlns:w="${W}" xmlns:w14="${W14}">
    <w:comment w:id="1" w:author="Reviewer" w:date="2026-10-02T00:00:00Z"><w:p w14:paraId="AAAA0001"><w:r><w:t>Root note.</w:t></w:r></w:p></w:comment>
    <w:comment w:id="2" w:author="Editor" w:date="2026-10-02T00:01:00Z"><w:p w14:paraId="CCCC0003"><w:r><w:t>Reply note.</w:t></w:r></w:p></w:comment>
    <w:comment w:id="3" w:author="Reviewer" w:date="2026-10-02T00:02:00Z"><w:p w14:paraId="BBBB0002"><w:r><w:t>Other root note.</w:t></w:r></w:p></w:comment>
  </w:comments>`;
  const commentsExtendedXml = `<w15:commentsEx xmlns:w15="${W15}">
    <w15:commentEx w15:paraId="AAAA0001" w15:done="0"/>
    <w15:commentEx w15:paraId="CCCC0003" w15:paraIdParent="AAAA0001" w15:done="0"/>
    <w15:commentEx w15:paraId="BBBB0002" w15:done="0"/>
  </w15:commentsEx>`;
  const commentsIdsXml = `<w16cid:commentsIds xmlns:w16cid="${W16CID}">
    <w16cid:commentId w16cid:paraId="AAAA0001" w16cid:durableId="11111111"/>
    <w16cid:commentId w16cid:paraId="CCCC0003" w16cid:durableId="22222222"/>
    <w16cid:commentId w16cid:paraId="BBBB0002" w16cid:durableId="33333333"/>
  </w16cid:commentsIds>`;
  let flatOpc = buildDocumentCommentsPackage(documentXml, commentsXml);
  flatOpc = addFlatOpcPart(flatOpc, 'commentsExtended.xml', commentsExtendedXml);
  return addFlatOpcPart(flatOpc, 'commentsIds.xml', commentsIdsXml);
}

let host = harness();
assert.deepEqual(await host.restore(), { success: true, verified: true, written: true, writeAttempted: true, message: 'Reverted and verified the body checkpoint.' });
assert.deepEqual(host.state(), { writes: 1, mode: 'TrackAll', modes: ['Off', 'TrackAll'] });
let receipt;

const threadedCheckpoint = buildThreadedCheckpoint();
const expectedThreadParts = inspectWordSource(threadedCheckpoint).parts;
const expectedThreadToken = await computeWordRestorationToken(expectedThreadParts);
const regenerateWordIds = xml => [
  ['10000001', '10000011'], ['10000002', '10000012'], ['10000003', '10000013'],
  ['AAAA0001', 'AAAA0011'], ['CCCC0003', 'CCCC0013'], ['BBBB0002', 'BBBB0012']
].reduce((result, [before, after]) => result.replaceAll(before, after), xml);
const regeneratedParts = inspectWordSource(regenerateWordIds(threadedCheckpoint)).parts;
const regeneratedToken = await computeWordRestorationToken(regeneratedParts);
assert.equal(regeneratedToken.value, expectedThreadToken.value);
assert.deepEqual(regeneratedToken.commentIdentity, expectedThreadToken.commentIdentity,
  'the fixture remaps body, comment, reply-parent and durable-link paragraph IDs while retaining thread relationships');
host = harness({ checkpoint: threadedCheckpoint, transform: regenerateWordIds });
receipt = await host.restore();
assert.equal(receipt.verified, true, 'Word-regenerated comment paragraph IDs verify when canonical reply/anchor identity is unchanged');
assert.equal(host.state().writes, 1);

// Word also regenerates the w14:paraId on table rows when a checkpoint is re-inserted (seen in Word, 2026-11-09).
const tableCheckpoint = buildDocumentFragmentPackage(
  `<w:tbl xmlns:w14="${W14}"><w:tr w14:paraId="4C32C4F9" w14:textId="77777777"><w:tc><w:p w14:paraId="20000001"><w:r><w:t>Cell text.</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>After the table.</w:t></w:r></w:p>`,
  { appendTrailingParagraph: false });
host = harness({ checkpoint: tableCheckpoint, transform: xml => xml.replace('4C32C4F9', '5F1D733B').replace('20000001', '20000011') });
receipt = await host.restore();
assert.equal(receipt.verified, true, 'a regenerated table-row id does not fail restore verification');
host = harness({ checkpoint: tableCheckpoint, transform: xml => xml.replace('Cell text.', 'Other text.') });
receipt = await host.restore();
assert.equal(receipt.verified, false, 'changed cell text still fails restore verification');
const tableGuard = async xml => (await computeWordRevertGuardToken(inspectWordSource(xml).parts)).value;
assert.equal(await tableGuard(tableCheckpoint.replace('4C32C4F9', '5F1D733B')), await tableGuard(tableCheckpoint),
  'the Revert guard ignores a regenerated table-row id');

// The same for a table inside a comment.
const commentTable = rowId => threadedCheckpoint.replace('<w:p w14:paraId="BBBB0002">',
  `<w:tbl><w:tr w14:paraId="${rowId}"><w:tc><w:p><w:r><w:t>Cell.</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="BBBB0002">`);
assert.equal((await computeWordRestorationToken(inspectWordSource(commentTable('5F1D733B')).parts)).value,
  (await computeWordRestorationToken(inspectWordSource(commentTable('4C32C4F9')).parts)).value,
  'a regenerated table-row id inside a comment keeps the restoration token');

const reparentedCheckpoint = threadedCheckpoint.replace(
  'w15:paraIdParent="AAAA0001"', 'w15:paraIdParent="BBBB0002"'
);
const reparentedParts = inspectWordSource(reparentedCheckpoint).parts;
const reparentedToken = await computeWordRestorationToken(reparentedParts);
assert.equal(reparentedToken.value, expectedThreadToken.value,
  'the full parts token alone intentionally omits regenerated parent paragraph IDs');
assert.notDeepEqual(reparentedToken.commentIdentity, expectedThreadToken.commentIdentity,
  'the canonical comment projection records a reply-parent change');
host = harness({ checkpoint: threadedCheckpoint, transform: () => reparentedCheckpoint });
receipt = await host.restore();
assert.equal(receipt.verified, false, 'restore requires both the full-parts token and canonical thread identity to match');
assert.equal(receipt.written, true);
assert.equal(host.state().writes, 1, 'a changed reply relationship is detected after the one replacement without replay');

const lostReplyCheckpoint = threadedCheckpoint.replace(/\s*<w:comment w:id="2"[^>]*>[\s\S]*?<\/w:comment>/, '');
host = harness({ checkpoint: lostReplyCheckpoint });
receipt = await host.restore();
assert.equal(receipt.writeAttempted, false, 'unmapped thread metadata is refused before any checkpoint write');
assert.equal(host.state().writes, 0);

host = harness({ failureSync: 3 }); // Content sync rejects after the one insertion.
receipt = await host.restore();
assert.equal(receipt.success, false);
assert.equal(receipt.writeAttempted, true);
assert.equal(receipt.written, false);
assert.equal(host.state().writes, 1, 'uncertain restoration must never replay');
assert.equal(host.state().mode, 'TrackAll');

host = harness({ transform: xml => xml.replace('Saved provision.', 'Different provision.') });
receipt = await host.restore();
assert.equal(receipt.verified, false);
assert.equal(receipt.written, true);
assert.equal(host.state().writes, 1);

host = harness({ postRunError: true });
receipt = await host.restore();
assert.equal(receipt.success, false);
assert.equal(receipt.written, true, 'post-callback errors must retain settled host evidence');

host = harness({ documentChanged: true });
receipt = await host.restore();
assert.equal(receipt.writeAttempted, false);
assert.equal(host.state().writes, 0);
assert.equal(host.state().mode, 'TrackAll');

host = harness();
receipt = await host.restore('');
assert.equal(receipt.writeAttempted, false);
assert.equal(host.state().writes, 0);
host = harness({ failureSync: 2 }); // Disabling tracking sync rejects before content insertion.
receipt = await host.restore();
assert.equal(receipt.writeAttempted, false);
assert.equal(host.state().writes, 0);
assert.equal(host.state().mode, 'TrackAll', 'even a failed tracking-disable sync must attempt to restore the prior mode');

const cancelled = new AbortController(); cancelled.abort();
host = harness({ controller: cancelled });
assert.equal((await host.restore()).writeAttempted, false);
assert.equal(host.state().writes, 0);
const duringCommit = new AbortController();
host = harness({ controller: duringCommit, cancelDuringCommit: true });
assert.equal((await host.restore()).verified, true, 'cancellation preserves settled restoration verification');
assert.equal(host.state().writes, 1);

console.log('PASS: checkpoint restoration verifies a single replacement, tracking restoration and uncertain-write evidence');

// Revert guard: nothing is written when the document changed after the checkpointed edit.
const afterEdit = buildDocumentFragmentPackage('<w:p><w:r><w:t>Edited provision.</w:t></w:r></w:p>', { appendTrailingParagraph: false });
const afterEditToken = (await computeWordRevertGuardToken(inspectWordSource(afterEdit).parts)).value;
host = harness({ currentXml: afterEdit, guardToken: afterEditToken });
receipt = await host.restore();
assert.equal(receipt.verified, true, 'an unchanged post-edit document reverts without asking');
assert.equal(host.state().writes, 1);

const laterTyping = buildDocumentFragmentPackage('<w:p><w:r><w:t>Edited provision, then typed by the user.</w:t></w:r></w:p>', { appendTrailingParagraph: false });
host = harness({ currentXml: laterTyping, guardToken: afterEditToken });
receipt = await host.restore();
const laterTypingToken = (await computeWordRevertGuardToken(inspectWordSource(laterTyping).parts)).value;
assert.deepEqual(receipt, { success: false, verified: false, written: false, writeAttempted: false, changedSinceEdit: true, currentToken: laterTypingToken, message: CHANGED_SINCE_EDIT_MESSAGE });
assert.deepEqual(host.state(), { writes: 0, mode: 'TrackAll', modes: [] }, 'a changed document is never replaced or retracked without confirmation');

// "Revert anyway" discards exactly the state the user was shown, and asks again about anything newer.
host = harness({ currentXml: laterTyping, guardToken: receipt.currentToken });
assert.equal((await host.restore()).verified, true, 'confirming discards the state shown in the question');
assert.equal(host.state().writes, 1);
const typedAgain = buildDocumentFragmentPackage('<w:p><w:r><w:t>Edited provision, then typed by the user, and more.</w:t></w:r></w:p>', { appendTrailingParagraph: false });
host = harness({ currentXml: typedAgain, guardToken: receipt.currentToken });
assert.equal((await host.restore()).changedSinceEdit, true, 'a stale confirmation never discards newer changes');
assert.equal(host.state().writes, 0);

// Word's background proofing after the edit (proofing marks, split runs) is not a change.
const proofed = buildDocumentFragmentPackage('<w:p><w:r><w:t xml:space="preserve">Edited </w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>provision.</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:lastRenderedPageBreak/></w:r></w:p>', { appendTrailingParagraph: false });
host = harness({ currentXml: proofed, guardToken: afterEditToken });
assert.equal((await host.restore()).verified, true, 'proofing noise alone reverts without asking');

host = harness({ currentXml: afterEdit, guardToken: null });
receipt = await host.restore();
assert.equal(receipt.changedSinceEdit, true, 'an unknown post-edit state is treated as changed');
assert.equal(host.state().writes, 0);

console.log('PASS: revert guard refuses to discard changes made after the checkpointed edit');
