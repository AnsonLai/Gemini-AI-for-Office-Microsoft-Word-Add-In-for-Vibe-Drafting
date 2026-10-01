# 0.8.2 standalone list operation gaps

**Package:** `@ansonlai/docx-redline-js@0.8.2`  
**Checked:** 2026-09-30  
**Scope:** public standalone document operations used by the add-in

## Plain header to list conversion

The standalone operation contract has no dedicated list conversion operation. `list-change` passes validation, but it is normalized as a redline and dispatched through the ordinary text mutation path. The list fallback in 0.8.2 converts a one-paragraph marker line only when its content is already equivalent to the source text. A plain paragraph cannot be converted by adding a marker in one canonical redline operation.

Minimal reproduction against the installed package:

```js
import './tests/setup-xml-provider.mjs';
import { inspectDocumentParts } from '@ansonlai/docx-redline-js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

const documentXml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Header</w:t></w:r></w:p><w:sectPr/></w:body></w:document>';
const result = await applyOperationsToDocumentXml(documentXml, [
  { type: 'redline', target: { index: 1, exactText: 'Header' }, modified: 'A. Header' }
], 'Planner test', {}, { atomic: true, strictTargets: true, generateRedlines: true });
const inspected = inspectDocumentParts({ documentXml: result.documentXml });

console.log(result.status, result.hasChanges, result.numberingXmlParts?.length || 0);
console.log(inspected.paragraphs[0].text, inspected.paragraphs[0].list);
```

Observed output: status `ok`, `hasChanges: true`, no numbering part, paragraph text `A. Header`, and `list: null`. This makes a redline appear successful while leaving the manual marker in ordinary text.

The fallback planner also refuses a changed heading payload. Its source and modified text must satisfy `sameRawText` or `sameListText`; for example, `1. Old heading` to `1. New heading` does not qualify. When a source already contains `1. Header`, passing an equivalent marker-prefixed modified value can convert the text-no-op into a real list paragraph.

`paragraph-format` does not provide a numbering escape hatch: the supported properties update paragraph alignment, keep settings, page-break-before, and style. They do not bind `w:numPr` or create numbering definitions. No application-side OOXML mutation is included in this report.

## Safe adapter behavior

Keep `convert_headers_to_list` on the existing host path for unmarked plain headers or changed replacement text. A canonical planner can handle only source marker lines whose requested text is unchanged, and only where the resulting numbering sequence has been verified against Word-authored fixtures. The planner should return an explicit unsupported-operation error for the other cases so the caller can preserve its established host behavior.

## List inspection and numbering fidelity

An initial diagnostic fixture used decimal/lower-letter definitions with literal bullet glyphs. The inspector reported those actual formats. The corrected frozen Word-authored fixture has `w:numFmt="bullet"` at every bullet level, and the inspector reports `list.format="bullet"`. No inspector defect is reproduced. Use the final source observations and package rather than the initial diagnostic's labels.
Applying the supported `edit_list` redline to bullet source paragraphs P3–P4 with `listType: 'bullet'` preserves the source `numId` and Word-authored glyphs at levels 0 and 1. The operation returns additional numbering XML, but `mergeNumberingXmlBySchemaOrder` skips incoming definitions whose IDs collide with existing ones. In this fixture the source `numId` therefore continues to use the source definitions. A caller cannot assume that requesting a different list kind or numbering style will replace an existing list definition while retaining the source `numId`. Format-changing edits need a dedicated upstream operation and Word-verified numbering remapping before migration.

The upper-letter manual-marker conversion case also emits xmldom hierarchy-parse warnings while generating list OOXML, despite returning a successful result. Treat this as an upstream diagnostic that needs a strict-parse and live Word round-trip check before broad migration; the offline package inspection alone is not sufficient evidence that every generated numbering part is valid.

Do not compensate for these behaviors by patching numbering XML in the add-in. Keep planner routing limited to operations whose source numbering instance is retained and whose output has been verified, and record unsupported list-kind/format changes explicitly.
