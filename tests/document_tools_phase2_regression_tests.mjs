import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import {
  acceptTrackedChangesInOoxml,
  inspectDocumentParts,
  rejectTrackedChangesInOoxml
} from '@ansonlai/docx-redline-js';
import { buildDocumentFragmentPackage, buildDocumentCommentsPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { inspectWordSource } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { createDocumentInspectionSession } from '../src/taskpane/modules/chat/document-inspection-session.js';
import { createDocumentTools } from '../src/taskpane/modules/chat/document-tools.js';

// Phase-2 WP3/WP4 regression coverage for the direct document tools:
// existingRevisions 'auto', post-commit host verification, and KB-7 signature lines.

globalThis.Word = {
  ChangeTrackingMode: { off: 'Off', trackAll: 'TrackAll' },
  InsertLocation: { replace: 'Replace', Replace: 'Replace' }
};

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const AUTHOR = 'Document Tools Test';

function fragment(paragraphXml, options = {}) {
  return buildDocumentFragmentPackage(paragraphXml, { appendTrailingParagraph: false, ...options });
}

function createHarness(initialXml, { generateRedlines = true, transformWrite } = {}) {
  const counts = { contentWrites: 0 };
  let currentXml = initialXml;
  const body = {
    getOoxml: () => ({ value: currentXml }),
    insertOoxml(payload) {
      counts.contentWrites++;
      currentXml = transformWrite ? transformWrite(payload) : payload;
    }
  };
  const document = { body, load() {}, changeTrackingMode: Word.ChangeTrackingMode.trackAll };
  const context = { document, async sync() {} };
  const session = createDocumentInspectionSession({
    readSource: async () => ({
      documentId: 'phase2-doc', revisionView: 'accepted', parts: inspectWordSource(currentXml).parts
    })
  });
  const tools = createDocumentTools({
    session,
    getDocumentId: () => 'phase2-doc',
    getAuthor: () => AUTHOR,
    getGenerateRedlines: () => generateRedlines,
    persistCheckpoint: async () => 1,
    signal: new AbortController().signal,
    run: async callback => callback(context)
  });
  return {
    tools, counts,
    get currentXml() { return currentXml; },
    async handleFor(needle) {
      const result = await tools.inspectDocument({ search: needle, limit: 8 });
      assert.equal(result.status, 'ok', JSON.stringify(result));
      const target = result.targets.find(item => item.text.includes(needle));
      assert.ok(target, `inspection returned a handle for ${JSON.stringify(needle)}`);
      return target.handle;
    }
  };
}

const paragraphTexts = (documentXml, parts) =>
  inspectDocumentParts({ ...parts, documentXml }, { revisionView: 'accepted' }).paragraphs.map(item => item.exactText);

function views(harness) {
  const { parts } = inspectWordSource(harness.currentXml);
  const accepted = acceptTrackedChangesInOoxml(parts.documentXml, { allAuthors: true });
  const rejected = rejectTrackedChangesInOoxml(parts.documentXml, { allAuthors: true });
  return {
    parts,
    accepted: paragraphTexts(accepted.oxml, parts),
    rejected: paragraphTexts(rejected.oxml, parts)
  };
}

const originalRejected = xml => {
  const { parts } = inspectWordSource(xml);
  return paragraphTexts(rejectTrackedChangesInOoxml(parts.documentXml, { allAuthors: true }).oxml, parts);
};

// ---------------------------------------------------------------------------
// existingRevisions 'auto': edits around another reviewer's or an unattributed pending insertion.
// ---------------------------------------------------------------------------
const insertionVariants = [
  { label: 'other reviewer', attrs: 'w:id="21" w:author="Prior Reviewer" w:date="2026-10-02T00:00:00Z"', author: 'Prior Reviewer' },
  { label: 'unattributed', attrs: 'w:id="22" w:author="" w:date="2026-10-02T00:00:00Z"', author: '' }
];
for (const variant of insertionVariants) {
  for (const kind of ['replace_text', 'rewrite_paragraph']) {
    const source = fragment(`<w:p><w:r><w:t xml:space="preserve">Base before </w:t></w:r><w:ins ${variant.attrs}><w:r><w:t>pending clause.</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> base after.</w:t></w:r></w:p>`);
    const harness = createHarness(source);
    const target = await harness.handleFor('pending clause.');
    const edit = kind === 'replace_text'
      ? { type: 'replace_text', target, find: 'Base before', replace: 'Revised before' }
      : { type: 'rewrite_paragraph', target, text: 'Revised before pending clause. base after.' };
    const receipt = await harness.tools.applyDocumentEdits({ edits: [edit] });
    const label = `${variant.label}/${kind}`;
    assert.equal(receipt.success, true, `${label}: ${JSON.stringify(receipt)}`);
    assert.equal(receipt.verified, true, label);
    assert.equal(harness.counts.contentWrites, 1, label);

    const { parts, accepted, rejected } = views(harness);
    assert.deepEqual(accepted, ['Revised before pending clause. base after.'], `${label}: Accept All gives the new text`);
    assert.deepEqual(rejected, originalRejected(source), `${label}: Reject All restores the original`);
    const insertions = parts.documentXml.match(/<w:ins\b[^>]*>/g) || [];
    const survivors = insertions.filter(tag => variant.author
      ? tag.includes(`w:author="${variant.author}"`) && tag.includes('w:id="21"')
      : /w:author=""/.test(tag));
    assert.ok(survivors.length >= 1, `${label}: the other author's w:ins survives (${insertions.join(' | ')})`);
    const otherText = [...parts.documentXml.matchAll(/<w:ins\b[^>]*w:author="Prior Reviewer"[^>]*>([\s\S]*?)<\/w:ins>/g)]
      .map(match => match[1]).join('');
    if (variant.author) assert.match(otherText, /pending clause\./, `${label}: the pending insertion keeps its text`);
    const ours = (parts.documentXml.match(/<w:(?:ins|del)\b[^>]*>/g) || []).filter(tag => tag.includes(`w:author="${AUTHOR}"`));
    assert.ok(ours.length > 0, `${label}: the edit is tracked under the configured author`);
  }
}

// Only the configured author's own pending revisions: merged, not stacked.
for (const kind of ['replace_text', 'rewrite_paragraph']) {
  const source = fragment(`<w:p><w:r><w:t xml:space="preserve">Base wording </w:t></w:r><w:ins w:id="23" w:author="${AUTHOR}" w:date="2026-10-02T00:00:00Z"><w:r><w:t>own pending change.</w:t></w:r></w:ins></w:p>`);
  const harness = createHarness(source);
  const target = await harness.handleFor('own pending change.');
  const edit = kind === 'replace_text'
    ? { type: 'replace_text', target, find: 'own pending change.', replace: 'own revised change.' }
    : { type: 'rewrite_paragraph', target, text: 'Base wording own revised change.' };
  const receipt = await harness.tools.applyDocumentEdits({ edits: [edit] });
  assert.equal(receipt.success, true, `own/${kind}: ${JSON.stringify(receipt)}`);
  assert.equal(receipt.verified, true);
  const { parts, accepted } = views(harness);
  assert.deepEqual(accepted, ['Base wording own revised change.'], `own/${kind}: Accept All gives the new text`);
  const xml = parts.documentXml;
  assert.doesNotMatch(xml, /<w:ins\b[^>]*>(?:(?!<\/w:ins>)[\s\S])*<w:(?:ins|del)\b/, `own/${kind}: no nested ins/del`);
  assert.doesNotMatch(xml, /<w:del\b[^>]*>(?:(?!<\/w:del>)[\s\S])*<w:ins\b/, `own/${kind}: no ins inside del`);
  assert.doesNotMatch(xml, /own pending change\./, `own/${kind}: the superseded own insertion text is merged away, not kept as a deletion`);
  const authors = [...new Set((xml.match(/<w:(?:ins|del)\b[^>]*w:author="([^"]*)"/g) || []).map(tag => tag.match(/w:author="([^"]*)"/)[1]))];
  assert.ok(authors.every(author => author === AUTHOR), `own/${kind}: only the configured author appears (${authors})`);
  assert.equal(originalRejected(source)[0], 'Base wording ');
  assert.equal(views(harness).rejected[0], 'Base wording ', `own/${kind}: Reject All still restores the base text`);
}

// ---------------------------------------------------------------------------
// Post-commit host verification.
// ---------------------------------------------------------------------------
async function runWithTransform({ source, edit, transformWrite, generateRedlines = false, needle }) {
  const harness = createHarness(source, { generateRedlines, transformWrite });
  const target = await harness.handleFor(needle);
  const receipt = await harness.tools.applyDocumentEdits({ edits: [{ ...edit, target }] });
  return { harness, receipt };
}
function assertHostMismatch(label, { harness, receipt }) {
  assert.equal(receipt.success, false, `${label}: ${JSON.stringify(receipt)}`);
  assert.equal(receipt.verified, false, label);
  assert.equal(receipt.error?.code, 'HOST_VERIFICATION_FAILED', label);
  assert.equal(receipt.written, true, label);
  assert.equal(receipt.requiresInspection, true, label);
  assert.equal(receipt.mutationOutcome, 'applied_with_host_error', label);
  assert.equal(harness.counts.contentWrites, 1, `${label}: never replayed`);
}
function assertHostMatch(label, { harness, receipt }) {
  assert.equal(receipt.success, true, `${label}: ${JSON.stringify(receipt)}`);
  assert.equal(receipt.verified, true, label);
  assert.equal(receipt.written, true, label);
  assert.equal(harness.counts.contentWrites, 1, label);
}

const plainSource = fragment('<w:p w14:paraId="1111AAAA" xmlns:w14="' + W14 + '"><w:r><w:t>Original clause text.</w:t></w:r></w:p>');
const replaceOriginal = { type: 'replace_text', find: 'Original', replace: 'Revised' };

// Text differs.
assertHostMismatch('text', await runWithTransform({
  source: plainSource, edit: replaceOriginal, needle: 'Original clause',
  transformWrite: xml => xml.replace('Revised', 'Rewritten')
}));

// Review segments differ although accepted text is unchanged (carrier flattened).
{
  const source = fragment(`<w:p><w:r><w:t xml:space="preserve">Base before </w:t></w:r><w:ins w:id="21" w:author="Prior Reviewer"><w:r><w:t>Pending clause.</w:t></w:r></w:ins></w:p>`);
  assertHostMismatch('review segments', await runWithTransform({
    source, edit: { type: 'replace_text', find: 'Base before', replace: 'Revised before' }, needle: 'Pending clause',
    transformWrite: xml => xml.replace(/<w:ins\b[^>]*>([\s\S]*?)<\/w:ins>/, '$1')
  }));
}

// List labels differ.
{
  const numberingXml = `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
  const source = fragment('<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Original item.</w:t></w:r></w:p>',
    { includeNumbering: true, numberingXml });
  const labelled = inspectWordSource(source).inspection.paragraphs[0].list;
  assert.equal(labelled?.label, '1.', 'fixture exposes a list label');
  let sawChange = false;
  const result = await runWithTransform({
    source, edit: replaceOriginal, needle: 'Original item',
    transformWrite: xml => {
      const changed = xml.replaceAll('w:val="%1."', 'w:val="%1)"');
      sawChange = changed !== xml;
      return changed;
    }
  });
  assert.equal(sawChange, true, 'the list-label transform applied to the prepared package');
  assertHostMismatch('list labels', result);
}

// Comments differ.
{
  const source = buildDocumentCommentsPackage(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Original clause.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p></w:body></w:document>`,
    `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="Reviewer"><w:p><w:r><w:t>Original comment.</w:t></w:r></w:p></w:comment></w:comments>`);
  let sawChange = false;
  const result = await runWithTransform({
    source, edit: replaceOriginal, needle: 'Original clause',
    transformWrite: xml => {
      const changed = xml.replace('Original comment.', 'Altered comment.');
      sawChange = changed !== xml;
      return changed;
    }
  });
  assert.equal(sawChange, true, 'the comment transform applied to the prepared package');
  assertHostMismatch('comments', result);
}

// Semantic-preserving host differences pass: ids, rsids, paragraph ids, run splits.
{
  const tracked = fragment(`<w:p><w:r><w:t xml:space="preserve">Base before </w:t></w:r><w:ins w:id="21" w:author="Prior Reviewer"><w:r><w:t>Pending clause.</w:t></w:r></w:ins></w:p>`);
  const trackedEdit = { type: 'replace_text', find: 'Base before', replace: 'Revised before' };
  const passes = [
    ['revision ids', tracked, trackedEdit, 'Pending clause', xml => xml.replace('w:id="21"', 'w:id="431"')],
    ['rsids', plainSource, replaceOriginal, 'Original clause',
      xml => xml.replace(/<w:p\b/, '<w:p w:rsidR="00ABCDEF" w:rsidRDefault="00ABCDEF"').replace(/<w:r>/g, '<w:r w:rsidR="00ABCDEF">')],
    ['paragraph ids', plainSource, replaceOriginal, 'Original clause',
      xml => xml.replace(/w14:paraId="[0-9A-F]{8}"/g, 'w14:paraId="7A7A7A7A"').replace(/<w:p\b(?![^>]*paraId)/, `<w:p xmlns:w14="${W14}" w14:paraId="7A7A7A7A"`)],
    ['run splits', plainSource, replaceOriginal, 'Original clause',
      xml => xml.replace('clause text.', 'clause </w:t></w:r><w:r><w:t xml:space="preserve">text.')]
  ];
  for (const [label, source, edit, needle, transformWrite] of passes) {
    let sawChange = false;
    const result = await runWithTransform({
      source, edit, needle, generateRedlines: false,
      transformWrite: xml => { const changed = transformWrite(xml); sawChange = changed !== xml; return changed; }
    });
    assert.equal(sawChange, true, `${label}: transform applied`);
    assertHostMatch(label, result);
  }
}

// ---------------------------------------------------------------------------
// KB-7: signature lines keep every underscore through a rewrite (tool level).
// ---------------------------------------------------------------------------
{
  const line = `Signed: ${'_'.repeat(25)}`;
  assert.equal(line.match(/_/g).length, 25);
  for (const generateRedlines of [true, false]) {
    const source = fragment(`<w:p><w:r><w:t xml:space="preserve">${line}</w:t></w:r></w:p>`);
    const harness = createHarness(source, { generateRedlines });
    const target = await harness.handleFor('Signed:');
    const receipt = await harness.tools.applyDocumentEdits({ edits: [{ type: 'rewrite_paragraph', target, text: `${line} (print name)` }] });
    assert.equal(receipt.success, true, `signature(redlines=${generateRedlines}): ${JSON.stringify(receipt)}`);
    assert.equal(receipt.verified, true);
    const { accepted } = views(harness);
    assert.deepEqual(accepted, [`${line} (print name)`], 'Accept All keeps all 25 underscores and the appended text');
    assert.equal(accepted[0].match(/_/g).length, 25);
  }
}

console.log('PASS: document tools existingRevisions auto, host verification tolerance/mismatch and signature-line (KB-7) regressions');
