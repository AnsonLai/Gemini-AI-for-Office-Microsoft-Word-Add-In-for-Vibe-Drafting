import assert from 'node:assert/strict';
import '../../tests/setup-xml-provider.mjs';
import {
  acceptTrackedChangesInOoxml,
  applyRedlineToOxml,
  inspectDocumentParts,
  rejectTrackedChangesInOoxml
} from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const originalText = 'This agreement is governed by local law.';
const modifiedText = 'This agreement is governed by applicable local law.';
const source = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111"><w:r><w:t>${originalText}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
const policy = 'slice-cross-author';
const author = 'Probe';
const baseOptions = {
  author,
  generateRedlines: false,
  existingRevisions: policy,
  structuredContent: false
};
const paragraph = inspectDocumentParts({ documentXml: source }).paragraphs[0];
const target = {
  index: paragraph.index,
  exactText: paragraph.exactText,
  paragraphId: paragraph.paragraphId,
  fingerprint: paragraph.fingerprint
};

function hasProbeInsertion(documentXml) {
  return /<w:ins\b[^>]*\bw:author="Probe"/.test(documentXml);
}

function publicViews(documentXml) {
  const accepted = acceptTrackedChangesInOoxml(documentXml, { allAuthors: true });
  const rejected = rejectTrackedChangesInOoxml(documentXml, { allAuthors: true });
  assert.equal(accepted.error, undefined, JSON.stringify(accepted.error));
  assert.equal(rejected.error, undefined, JSON.stringify(rejected.error));
  return {
    accepted: inspectDocumentParts({ documentXml: accepted.oxml }).paragraphs[0].exactText,
    rejected: inspectDocumentParts({ documentXml: rejected.oxml }).paragraphs[0].exactText
  };
}

const direct = await applyRedlineToOxml(source, originalText, modifiedText, baseOptions);
const directViews = publicViews(direct.oxml);
assert.equal(direct.status, 'ok', JSON.stringify(direct.error));
assert.equal(direct.hasChanges, true);
assert.equal(directViews.accepted, modifiedText);
assert.equal(directViews.rejected, originalText);

const nonSurgicalControl = await applyRedlineToOxml(source, originalText, modifiedText, {
  ...baseOptions,
  existingRevisions: 'merge-same-author'
});
const controlViews = publicViews(nonSurgicalControl.oxml);
assert.equal(nonSurgicalControl.status, 'ok', JSON.stringify(nonSurgicalControl.error));
assert.equal(nonSurgicalControl.hasChanges, true);
assert.equal(hasProbeInsertion(nonSurgicalControl.oxml), false, 'the non-surgical control is untracked');
assert.equal(controlViews.accepted, modifiedText);
assert.equal(controlViews.rejected, modifiedText);

const batchCases = [
  { label: 'redline.modified', operation: { type: 'redline', target, modified: modifiedText } },
  { label: 'redline.replacements', operation: { type: 'redline', target, replacements: [{ find: 'governed by local law', replace: 'governed by applicable local law' }] } },
  { label: 'replace.modified', operation: { type: 'replace', target, modified: modifiedText } },
  { label: 'replace.replacements', operation: { type: 'replace', target, replacements: [{ find: 'governed by local law', replace: 'governed by applicable local law' }] } }
];

const batchResults = [];
for (const { label, operation } of batchCases) {
  const result = await applyOperationsToDocumentXml(
    source,
    [{ ...operation, structuredContent: false, generateRedlines: false, existingRevisions: policy }],
    author,
    null,
    { generateRedlines: false, existingRevisions: policy }
  );
  assert.equal(result.status, 'ok', `${label}: ${JSON.stringify(result.results?.[0]?.error)}`);
  assert.equal(result.hasChanges, true, label);
  const views = publicViews(result.documentXml);
  assert.equal(views.accepted, modifiedText, `${label} accepted view`);
  assert.equal(views.rejected, originalText, `${label} rejected view`);
  batchResults.push({ label, hasProbeInsertion: hasProbeInsertion(result.documentXml), ...views });
}

const observed = hasProbeInsertion(direct.oxml)
  && batchResults.every(result => result.hasProbeInsertion)
  && directViews.accepted === modifiedText
  && directViews.rejected === originalText;

if (!observed) {
  console.error(JSON.stringify({
    issue: 'generateRedlines false + slice-cross-author insertion no longer reproduces',
    directHasProbeInsertion: hasProbeInsertion(direct.oxml),
    directViews,
    controlHasProbeInsertion: hasProbeInsertion(nonSurgicalControl.oxml),
    controlViews,
    batchResults
  }, null, 2));
  process.exitCode = 1;
} else {
  console.log(JSON.stringify({
    package: '@ansonlai/docx-redline-js@0.8.4',
    reproduced: true,
    direct: { hasProbeInsertion: true, ...directViews },
    mergeSameAuthorControl: { hasProbeInsertion: false, ...controlViews },
    batchResults
  }, null, 2));
}
