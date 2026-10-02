// Reproduces (0.8.4 regression): two redline operations in ONE batch that turn
// non-adjacent manually numbered headers into "A." / "B." list items now get
// two different numbering instances, so the second restarts at "A.".
// docx-redline-js 0.8.3 gave both the same numId ("A.", "B.").
// Exits 0 while the defect reproduces.
//
//   node docs/library-issues/2026-10-02-separate-list-operations-restart-numbering-reproducer.mjs
import '../../tests/setup-xml-provider.mjs';
import { readFileSync } from 'node:fs';
import { mergeNumberingXmlBySchemaOrder, openDocx } from '@ansonlai/docx-redline-js';
import { unzipDocx, zipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';

// Word-authored fixture with "1. Supported Header First", a body paragraph, and
// "2. Supported Header Second" (paragraphs 21-23).
const bytes = new Uint8Array(readFileSync(new URL('../../tests/fixtures/agentic-lists/nested-lists-source.docx', import.meta.url)));
const source = openDocx(bytes);
const decode = name => new TextDecoder().decode(source.entries.get(name));
const paragraphs = source.inspect().paragraphs;
const numberingXml = decode('word/numbering.xml');
const target = text => {
  const paragraph = paragraphs.find(item => item.exactText === text);
  return { index: paragraph.index, exactText: paragraph.exactText, paragraphId: paragraph.paragraphId, fingerprint: paragraph.fingerprint };
};

const result = await applyOperationsToDocumentXml(decode('word/document.xml'), [
  { type: 'redline', target: target('1. Supported Header First'), modified: 'A. Supported Header First' },
  { type: 'redline', target: target('2. Supported Header Second'), modified: 'B. Supported Header Second' }
], 'Reproducer', { numberingXml, stylesXml: decode('word/styles.xml') },
{ atomic: true, strictTargets: true, generateRedlines: true, structuredContent: true, existingRevisions: 'merge-same-author' });

let merged = numberingXml;
for (const part of result.numberingXmlParts || []) merged = mergeNumberingXmlBySchemaOrder(merged, part);
const entries = unzipDocx(bytes);
entries.set('word/document.xml', new TextEncoder().encode(result.documentXml));
entries.set('word/numbering.xml', new TextEncoder().encode(merged));
const accepted = openDocx(zipDocx(entries));
await accepted.resolveRevisions('accept', { allAuthors: true });
const list = text => accepted.inspect().paragraphs.find(item => item.exactText === text)?.list ?? null;
const first = list('Supported Header First');
const second = list('Supported Header Second');

const reproduced = result.status === 'ok' && first && second && first.numId !== second.numId;
console.log(`status=${result.status}; First=${JSON.stringify(first)}; Second=${JSON.stringify(second)}`);
console.log(reproduced
  ? 'REPRODUCED: the headers land in different lists, so the second restarts at "A.".'
  : 'NOT REPRODUCED: both headers share one list.');
process.exit(reproduced ? 0 : 1);
