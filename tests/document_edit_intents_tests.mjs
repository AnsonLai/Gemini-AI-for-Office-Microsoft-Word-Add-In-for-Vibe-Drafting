import './setup-xml-provider.mjs';
import assert from 'node:assert/strict';
import {
  acceptTrackedChangesInOoxml,
  inspectDocumentParts
} from '@ansonlai/docx-redline-js';
import { buildDocumentFragmentPackage } from '@ansonlai/docx-redline-js/services/package-builder.js';
import { prepareCanonicalBatch, captureSourceBaseline } from '../src/taskpane/modules/docx-redline-js-integration/consumer-core.js';
import {
  compileDocumentEdits,
  DocumentEditIntentError
} from '../src/taskpane/modules/chat/document-edit-intents.js';

function makePackage(paragraphAndTableXml) {
  return buildDocumentFragmentPackage(paragraphAndTableXml, { appendTrailingParagraph: false });
}

function targetsFrom(packageXml) {
  return captureSourceBaseline(packageXml).map((paragraph, index) => ({
    ...paragraph,
    handle: `handle-${index + 1}`,
    location: `Paragraph ${index + 1}`,
    fullTextInspected: true,
    kind: 'paragraph'
  }));
}

function createSession(targets, { endHandle = null } = {}) {
  const byHandle = new Map(targets.map(target => [target.handle, target]));
  if (endHandle) byHandle.set(endHandle.handle, endHandle);
  const calls = [];
  return {
    calls,
    async resolveHandles(handles, options) {
      calls.push({ handles: [...handles], options });
      return {
        status: 'ok',
        source: options.source,
        targets: handles.map(handle => {
          const target = byHandle.get(handle);
          if (!target) return null;
          const { handle: resolvedHandle, location, fullTextInspected, kind, position, ...descriptor } = target;
          return {
            handle: resolvedHandle,
            location,
            fullTextInspected,
            kind,
            ...(position ? { position } : {}),
            descriptor
          };
        })
      };
    }
  };
}

async function compileFor(packageXml, edits) {
  const targets = targetsFrom(packageXml);
  const source = { documentId: 'doc-1', revision: 4 };
  const session = createSession(targets);
  const compiled = await compileDocumentEdits(edits, { session, source });
  assert.equal(session.calls.length, 1, 'all opaque handles resolve against one fresh snapshot');
  assert.deepEqual(session.calls[0].options, { source });
  return { compiled, targets, session, source };
}

const sourcePackage = makePackage([
  '<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>A. Duties and C:\\Temp.</w:t></w:r></w:p>',
  '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
  '<w:p><w:r><w:t>Policy text here.</w:t></w:r></w:p>',
  '<w:p><w:r><w:t>Formatting sample.</w:t></w:r></w:p>'
].join(''));
const sourceParagraphs = targetsFrom(sourcePackage);
assert.equal(sourceParagraphs.length, 5, 'fixture exposes body and both table-cell paragraphs');
assert.equal(sourceParagraphs[1].inTable, true);
assert.equal(sourceParagraphs[2].inTable, true);

const source = { documentId: 'doc-1', revision: 4 };
const session = createSession(sourceParagraphs);
const compiled = await compileDocumentEdits([
  { type: 'replace_text', target: 'handle-1', find: 'A. Duties', replace: 'A. Responsibilities' },
  { type: 'rewrite_paragraph', target: 'handle-2', text: 'Cell A revised' },
  { type: 'comment', target: 'handle-4', anchor: 'Policy text', text: 'Confirm this policy.' },
  { type: 'rewrite_paragraph', target: 'handle-4', text: 'Policy text revised.' },
  { type: 'format_text', target: 'handle-5', find: 'Formatting', properties: { bold: true, highlight: 'yellow' } }
], { session, source });
assert.equal(session.calls.length, 1);
assert.deepEqual(session.calls[0].handles, ['handle-1', 'handle-2', 'handle-4', 'handle-5']);
assert.equal(compiled.requestedCount, 5);
assert.equal(compiled.operations[0].type, 'replace');
assert.equal(compiled.operations[0].structuredContent, false, 'literal replacements never enter structured/Markdown parsing');
assert.equal(compiled.operations[0].replacements[0].replace, 'A. Responsibilities');
assert.equal(compiled.operations[1].target.inTable, true, 'canonical table-cell target is retained');
assert.deepEqual(compiled.operations[2], {
  type: 'comment',
  target: compiled.operations[2].target,
  textToComment: 'Policy text',
  commentContent: 'Confirm this policy.'
});
assert.equal(compiled.operations[3].structuredContent, false);
assert.deepEqual(compiled.operations[4].properties, { bold: true, highlight: 'yellow' });
assert.equal(compiled.expectations.length, 5, 'every intent receives an expectation');
assert.equal(compiled.locations.length, 5, 'every intent retains its inspected location');

const prepared = await prepareCanonicalBatch(sourcePackage, compiled.operations, {
  author: 'Intent Compiler Test', existingRevisions: 'merge-same-author'
});
assert.equal(prepared.status, 'ready', JSON.stringify(prepared.result?.error || prepared.result?.results));
assert.ok(prepared.result.commentsXml, 'the library batch emits the requested comment part');
assert.match(prepared.result.commentsXml, /Confirm this policy\./);
const acceptedResult = acceptTrackedChangesInOoxml(prepared.result.documentXml, { allAuthors: true });
assert.equal(acceptedResult.status, undefined, JSON.stringify(acceptedResult.error));
const after = inspectDocumentParts({ documentXml: acceptedResult.oxml });
assert.deepEqual(after.paragraphs.filter(paragraph => paragraph.inTable).map(paragraph => paragraph.exactText), [
  'Cell A revised', 'Cell B'
], 'rewriting one table cell preserves the other cell and the table structure');
assert.equal(after.paragraphs[0].exactText, 'A. Responsibilities and C:\\Temp.');
assert.equal(after.paragraphs[3].exactText, 'Policy text revised.');
assert.match(acceptedResult.oxml, /<w:b(?:\s|\/|>)/, 'basic emphasis is generated by the canonical library');
assert.match(acceptedResult.oxml, /<w:i(?:\s|\/|>)/, 'localized text replacement preserves existing character formatting');
assert.match(acceptedResult.oxml, /w:val="yellow"/, 'yellow highlight is supported through the narrow format property');

const backslashPackage = makePackage('<w:p><w:r><w:t>Use A/B at C:\\Work\\Docs.</w:t></w:r></w:p>');
const backslashCompiled = await compileFor(backslashPackage, [
  { type: 'replace_text', target: 'handle-1', find: 'A/B', replace: 'A\\B' }
]);
assert.equal(backslashCompiled.compiled.operations[0].structuredContent, false);
const backslashPrepared = await prepareCanonicalBatch(backslashPackage, backslashCompiled.compiled.operations, {
  author: 'Intent Compiler Test'
});
assert.equal(backslashPrepared.status, 'ready', JSON.stringify(backslashPrepared.result?.error));
const backslashAccepted = acceptTrackedChangesInOoxml(backslashPrepared.result.documentXml, { allAuthors: true });
assert.equal(inspectDocumentParts({ documentXml: backslashAccepted.oxml }).paragraphs[0].exactText, 'Use A\\B at C:\\Work\\Docs.',
  'literal slash and backslash characters survive the edit without decoding');

const nbspPackage = makePackage('<w:p><w:r><w:t>Use A&#xA0;B here.</w:t></w:r></w:p>');
const nbspCompiled = await compileFor(nbspPackage, [
  { type: 'replace_text', target: 'handle-1', find: 'A B', replace: 'A C' }
]);
const nbspPrepared = await prepareCanonicalBatch(nbspPackage, nbspCompiled.compiled.operations, {
  author: 'Intent Compiler Test'
});
assert.equal(nbspPrepared.status, 'ready', JSON.stringify(nbspPrepared.result?.error));
const nbspAccepted = acceptTrackedChangesInOoxml(nbspPrepared.result.documentXml, { allAuthors: true });
assert.equal(inspectDocumentParts({ documentXml: nbspAccepted.oxml }).paragraphs[0].exactText, 'Use A\u00a0C here.',
  'the library owns narrow ASCII-space/NBSP matching and preserves the source spacing');

const ambiguousPackage = makePackage('<w:p><w:r><w:t>Repeat once, then Repeat again.</w:t></w:r></w:p>');
const ambiguousReplacement = await compileFor(ambiguousPackage, [
  { type: 'replace_text', target: 'handle-1', find: 'Repeat', replace: 'Change' }
]);
assert.equal(ambiguousReplacement.compiled.operations.length, 1,
  'the consumer forwards ambiguous literal requests without implementing a parallel matcher');
const ambiguousPrepared = await prepareCanonicalBatch(ambiguousPackage, ambiguousReplacement.compiled.operations);
assert.equal(ambiguousPrepared.status, 'refused');
assert.equal(ambiguousPrepared.result.results[0].error.code, 'AMBIGUOUS_PATCH_SOURCE',
  'the library reports ambiguity for the atomic batch');
assert.equal(ambiguousPrepared.insertionPayload, undefined, 'ambiguous source refuses before host insertion');

const selectedOccurrence = await compileFor(ambiguousPackage, [
  { type: 'replace_text', target: 'handle-1', find: 'Repeat', replace: 'Change', occurrence: 2 }
]);
const selectedOccurrencePrepared = await prepareCanonicalBatch(ambiguousPackage, selectedOccurrence.compiled.operations);
assert.equal(selectedOccurrencePrepared.status, 'ready', JSON.stringify(selectedOccurrencePrepared.result?.error));
const selectedOccurrenceAccepted = acceptTrackedChangesInOoxml(selectedOccurrencePrepared.result.documentXml, { allAuthors: true });
assert.equal(inspectDocumentParts({ documentXml: selectedOccurrenceAccepted.oxml }).paragraphs[0].exactText,
  'Repeat once, then Change again.', 'the library applies the requested 1-based occurrence');

const missingReplacement = await compileFor(ambiguousPackage, [
  { type: 'replace_text', target: 'handle-1', find: 'Missing', replace: 'Change' }
]);
const missingPrepared = await prepareCanonicalBatch(ambiguousPackage, missingReplacement.compiled.operations);
assert.equal(missingPrepared.status, 'refused');
assert.equal(missingPrepared.result.results[0].error.code, 'PATCH_SOURCE_NOT_FOUND',
  'the library, not a consumer-side count, owns missing-source diagnostics');

const insertPackage = makePackage('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Existing clause.</w:t></w:r></w:p>');
const insertCompiled = await compileFor(insertPackage, [
  { type: 'insert_paragraph', target: 'handle-1', position: 'after', text: 'Additional plain paragraph.' }
]);
assert.equal(insertCompiled.compiled.operations[0].structuredContent, true,
  'structured parsing is enabled only for the explicitly validated paragraph insertion');
assert.equal(insertCompiled.compiled.operations[0].modified, 'Existing clause.\nAdditional plain paragraph.');
const insertPrepared = await prepareCanonicalBatch(insertPackage, insertCompiled.compiled.operations, {
  author: 'Intent Compiler Test'
});
assert.equal(insertPrepared.status, 'ready', JSON.stringify(insertPrepared.result?.error));
const insertAccepted = acceptTrackedChangesInOoxml(insertPrepared.result.documentXml, { allAuthors: true });
assert.deepEqual(inspectDocumentParts({ documentXml: insertAccepted.oxml }).paragraphs.map(paragraph => paragraph.exactText), [
  'Existing clause.', 'Additional plain paragraph.'
]);
assert.match(insertAccepted.oxml, /<w:b(?:\s|\/|>)/, 'plain paragraph insertion retains existing paragraph character formatting');

const formatCommentPackage = makePackage('<w:p><w:r><w:t>Format and comment this.</w:t></w:r></w:p>');
const formatComment = await compileFor(formatCommentPackage, [
  { type: 'format_text', target: 'handle-1', find: 'Format', properties: { bold: true } },
  { type: 'comment', target: 'handle-1', anchor: 'Format', text: 'Review this wording.' }
]);
const formatCommentPrepared = await prepareCanonicalBatch(formatCommentPackage, formatComment.compiled.operations, {
  author: 'Intent Compiler Test'
});
assert.equal(formatCommentPrepared.status, 'ready', JSON.stringify(formatCommentPrepared.result?.error));
assert.match(formatCommentPrepared.result.commentsXml, /Review this wording\./);
assert.match(formatCommentPrepared.result.documentXml, /<w:b(?:\s|\/|>)/, 'format and comment on one source paragraph combine in one library batch');

const replacementCommentSession = createSession(targetsFrom(formatCommentPackage));
await assert.rejects(
  () => compileDocumentEdits([
    { type: 'replace_text', target: 'handle-1', find: 'Format', replace: 'Change' },
    { type: 'comment', target: 'handle-1', anchor: 'comment', text: 'Review this sentence.' }
  ], { session: replacementCommentSession, source }),
  error => error instanceof DocumentEditIntentError
    && error.code === 'COMMENT_ANCHOR_REQUIRES_REINSPECTION'
    && error.recovery.requiresReinspection,
  'literal replacement plus comment refuses even when consumer-computed ranges appear disjoint');

for (const proposedText of ['A replacement without the original anchor.', 'Policy text and Policy text again.']) {
  const rewriteCommentSession = createSession(sourceParagraphs);
  await assert.rejects(
    () => compileDocumentEdits([
      { type: 'rewrite_paragraph', target: 'handle-4', text: proposedText },
      { type: 'comment', target: 'handle-4', anchor: 'Policy text', text: 'Review this wording.' }
    ], { session: rewriteCommentSession, source }),
    error => error instanceof DocumentEditIntentError
      && error.code === 'COMMENT_ANCHOR_REQUIRES_REINSPECTION'
      && error.recovery.requiresReinspection,
    'rewrite plus comment requires one unchanged, unique original anchor in the proposed literal text');
}

const atomicPackage = makePackage([
  '<w:p><w:r><w:t>First stable paragraph.</w:t></w:r></w:p>',
  '<w:p><w:r><w:t>Second stable paragraph.</w:t></w:r></w:p>'
].join(''));
const atomicBatch = await compileFor(atomicPackage, [
  { type: 'rewrite_paragraph', target: 'handle-1', text: 'First proposed change.' },
  { type: 'format_text', target: 'handle-2', find: 'missing anchor', properties: { italic: true } }
]);
const atomicPrepared = await prepareCanonicalBatch(atomicPackage, atomicBatch.compiled.operations, {
  author: 'Intent Compiler Test'
});
assert.equal(atomicPrepared.status, 'refused', 'an engine-invalid member refuses the entire atomic batch');
assert.equal(atomicPrepared.insertionPayload, undefined, 'a refused batch cannot reach the host insertion');

const noopCompiled = await compileFor(insertPackage, [
  { type: 'rewrite_paragraph', target: 'handle-1', text: 'Existing clause.' }
]);
const noopPrepared = await prepareCanonicalBatch(insertPackage, noopCompiled.compiled.operations, {
  author: 'Intent Compiler Test'
});
assert.equal(noopPrepared.status, 'noop', 'a same-text rewrite remains an explicit canonical no-op');
assert.equal(noopPrepared.result.hasChanges, false);

const identitySession = createSession(sourceParagraphs);
await assert.rejects(
  () => compileDocumentEdits([
    { type: 'replace_text', target: 'handle-1', find: 'A. Duties', replace: 'A. Duties' }
  ], { session: identitySession, source }),
  error => error instanceof DocumentEditIntentError
    && error.code === 'UNSUPPORTED_LITERAL_NOOP'
    && /rewrite_paragraph/.test(error.message),
  'the library explicitly rejects identity literal patches; use complete desired-state rewrite for a valid no-op');
assert.equal(identitySession.calls.length, 0,
  'an unsupported identity literal patch refuses before target resolution or library preparation');

const alreadyBoldPackage = makePackage('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Already bold.</w:t></w:r></w:p>');
const alreadyBold = await compileFor(alreadyBoldPackage, [
  { type: 'format_text', target: 'handle-1', find: 'Already', properties: { bold: true } }
]);
const alreadyBoldPrepared = await prepareCanonicalBatch(alreadyBoldPackage, alreadyBold.compiled.operations);
assert.equal(alreadyBoldPrepared.status, 'noop', 'already-applied formatting remains a canonical no-op');

const highlightedPackage = makePackage('<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Marked.</w:t></w:r></w:p>');
const removeHighlight = await compileFor(highlightedPackage, [
  { type: 'format_text', target: 'handle-1', find: 'Marked.', properties: { highlight: 'none' } }
]);
const removeHighlightPrepared = await prepareCanonicalBatch(highlightedPackage, removeHighlight.compiled.operations);
assert.equal(removeHighlightPrepared.status, 'ready', JSON.stringify(removeHighlightPrepared.result?.error));
const unhighlighted = acceptTrackedChangesInOoxml(removeHighlightPrepared.result.documentXml, { allAuthors: true });
assert.doesNotMatch(unhighlighted.oxml, /w:highlight/, 'the narrow none value removes existing highlighting');

const appendPackage = makePackage('<w:p><w:r><w:t>Last paragraph.</w:t></w:r></w:p>');
const appendTarget = { ...targetsFrom(appendPackage)[0], handle: 'end-handle', kind: 'document-end', position: 'after', fullTextInspected: false };
const appendSession = createSession([], { endHandle: appendTarget });
const appendCompiled = await compileDocumentEdits([
  { type: 'append_paragraph', target: 'end-handle', text: 'Final plain paragraph.' }
], { session: appendSession, source });
assert.equal(appendCompiled.operations[0].structuredContent, true);
assert.equal(appendCompiled.operations[0].modified, 'Last paragraph.\nFinal plain paragraph.');
const appendPrepared = await prepareCanonicalBatch(appendPackage, appendCompiled.operations);
assert.equal(appendPrepared.status, 'ready', JSON.stringify(appendPrepared.result?.error));
const appendAccepted = acceptTrackedChangesInOoxml(appendPrepared.result.documentXml, { allAuthors: true });
assert.deepEqual(inspectDocumentParts({ documentXml: appendAccepted.oxml }).paragraphs.map(paragraph => paragraph.exactText), [
  'Last paragraph.', 'Final plain paragraph.'
]);

const incompleteTargets = targetsFrom(sourcePackage).map(target => ({ ...target }));
incompleteTargets[0].fullTextInspected = false;
const incompleteSession = createSession(incompleteTargets);
await assert.rejects(
  () => compileDocumentEdits([{ type: 'rewrite_paragraph', target: 'handle-1', text: 'Replacement.' }], {
    session: incompleteSession, source
  }),
  error => error instanceof DocumentEditIntentError
    && error.code === 'TARGET_TEXT_NOT_FULLY_INSPECTED'
    && error.recovery.requiresReinspection
);

const invalidSession = createSession(sourceParagraphs);
await assert.rejects(
  () => compileDocumentEdits([
    { type: 'rewrite_paragraph', target: 'handle-1', text: 'Would otherwise be valid.' },
    { type: 'format_text', target: 'handle-4', find: 'Policy', properties: { color: 'red' } }
  ], { session: invalidSession, source }),
  error => error instanceof DocumentEditIntentError && error.code === 'UNSUPPORTED_DOCUMENT_EDIT'
);
assert.equal(invalidSession.calls.length, 0, 'invalid mixed batches refuse before target resolution or library preparation');

const conflicts = createSession(sourceParagraphs);
await assert.rejects(
  () => compileDocumentEdits([
    { type: 'rewrite_paragraph', target: 'handle-1', text: 'First result.' },
    { type: 'delete_paragraph', target: 'handle-1' }
  ], { session: conflicts, source }),
  error => error instanceof DocumentEditIntentError && error.code === 'CONFLICTING_DOCUMENT_EDITS'
);

const newAnchorSession = createSession(sourceParagraphs);
await assert.rejects(
  () => compileDocumentEdits([
    { type: 'rewrite_paragraph', target: 'handle-4', text: 'Fresh text.' },
    { type: 'comment', target: 'handle-4', anchor: 'Fresh text', text: 'New comment.' }
  ], { session: newAnchorSession, source }),
  error => error instanceof DocumentEditIntentError
    && error.code === 'COMMENT_ANCHOR_REQUIRES_REINSPECTION'
    && error.recovery.requiresReinspection
);

console.log('PASS: typed document edit intent compiler tests');
