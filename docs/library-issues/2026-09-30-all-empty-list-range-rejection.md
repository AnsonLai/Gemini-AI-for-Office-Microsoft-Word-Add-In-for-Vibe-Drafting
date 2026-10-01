# All-empty list ranges lose their source paragraphs on Reject All

**Resolution:** Fixed in installed 0.8.3. The reproduction below describes the earlier defect; see [0.8.3 validation](../validation-reports/2026-09-30-docx-redline-v083.md).

**Checked:** 2026-09-30, unpublished library working tree at
`C:/Users/Phara/Desktop/Projects/Docx Redline JS`, still versioned 0.8.2.
**Evidence:** independently reproduced offline; no live Word claim for this case.

The new empty-middle-item test passes for `Alpha`, empty, `Gamma`, but replacing
a range containing only empty list paragraphs still returns `ok` and loses the
source paragraphs on rejection. This is a separate upstream follow-up; the
add-in has no library reconstruction workaround.

```js
import { applyOperationsToDocumentXml } from './services/standalone-operation-runner.js';
import { inspectDocumentParts, rejectTrackedChangesInOoxml } from './index.js';

const item = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>';
const xml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Intro</w:t></w:r></w:p>'
  + item + item + '<w:sectPr/></w:body></w:document>';
const result = await applyOperationsToDocumentXml(xml, [{
  type: 'redline', target: { index: 2, exactText: '' },
  targetEnd: { index: 3, exactText: '' }, modified: '- One\n- Two'
}], 'Review', {}, {
  atomic: true, strictTargets: true, structuredContent: true, generateRedlines: true
});
const rejected = rejectTrackedChangesInOoxml(result.documentXml, { allAuthors: true }).oxml;
console.log(result.status, inspectDocumentParts({ documentXml: rejected }).paragraphs.map(p => p.exactText));
```

Expected: `ok`, `['Intro', '', '']`. Actual: `ok`, `['Intro']`.
The receipt contains only insertion/structural revisions, without deleted source
paragraph marks. In `pipeline/list-generation.js`, source-paragraph preservation
is guarded by `deletionRuns.length > 0`; empty-only ranges have no text deletion
runs. Preserve their marks independently of text, or refuse before mutation until
supported. Add all-empty and single-empty cases alongside the mixed-empty test,
then verify Word Accept All/Reject All and exact paragraph counts.
