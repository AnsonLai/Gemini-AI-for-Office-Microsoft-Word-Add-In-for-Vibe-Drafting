# 0.8.2 historical list properties reported as active

**Resolution:** Fixed in installed 0.8.3. The reproduction below describes the earlier defect; see [0.8.3 validation](../validation-reports/2026-09-30-docx-redline-v083.md).

**Package:** `@ansonlai/docx-redline-js@0.8.2`  
**Checked:** 2026-09-30  
**Evidence:** Offline reproduction using the installed package and the frozen Word-authored list fixture; no live Word verification is claimed.

## Minimal inspection reproduction

`w:pPrChange` stores the paragraph properties that were replaced by a tracked formatting change. In this example, the paragraph's current properties contain only the change record; the `w:numPr` exists in the historical `w:pPr` nested under that record. The second paragraph is a control with an active `w:numPr` directly under its current `w:pPr`.

```xml
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pPrChange w:id="7" w:author="Prior Editor">
          <w:pPr>
            <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
          </w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>Plain paragraph with historical list properties</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>Active list paragraph</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>
```

Pass the document XML to the package's `inspectDocumentParts({ documentXml })`. Both paragraphs receive a non-null `list` with `numId: "1"` and `level: 0`. In the production fixture, which also supplies the Word-authored numbering part, the historical paragraph is reported as `{ numId: "1", level: 0, label: "", format: "bullet" }`.

## Source and operation evidence

In the installed package, `services/document-inspection.js:9-10` defines `descendants` and `first` using `getElementsByTagNameNS`, which searches descendants. `paragraphListProperties` at lines 20-26 asks for the first `numPr` descendant of the first `pPr` descendant of the paragraph. That lookup therefore crosses the `pPrChange` boundary and reads the old properties as though they were current. The function is used to populate the inspection `list` field at line 285.

An offline production-route test adds this historical subtree to a plain paragraph from `tests/fixtures/agentic-lists/nested-lists-source.docx`. The package then classifies it as an active bullet list. The list operation refuses before any Word write: the underlying library operation result is `EXISTING_REVISIONS` because the target contains a tracked formatting change by another author. The batch wrapper reports `BATCH_OPERATION_FAILED`; the consumer observer currently preserves that wrapper error in `operationResults` and returns a refused receipt, but does not include the nested library `results[0].error` detail. There is no `body.insertOoxml` call and no native insertion replay.

This confirms an offline inspection boundary defect and the current fail-closed behavior for this test case. It does not establish how every Word-authored `pPrChange` variation is serialized or inspected in a live Word host.

## Consumer behavior

Do not infer that the paragraph is a current list item from this inspection result. Until the library distinguishes current properties from historical `pPrChange` properties, a canonical operation that encounters this case must refuse safely and must not replay a native edit based on the same ambiguous classification. The ordinary plain-paragraph native insertion path remains covered separately. No consumer-side XML parser or historical-property workaround is included here.
