import '../setup-xml-provider.mjs';
import assert from 'assert';
import crypto from 'crypto';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

import { NS_W, resetRevisionIdCounter } from '../../src/taskpane/modules/reconciliation/core/types.js';
import { ingestOoxml } from '../../src/taskpane/modules/reconciliation/pipeline/ingestion.js';
import { applyRedlineToOxml } from '../../src/taskpane/modules/reconciliation/engine/oxml-engine.js';
import { generateTableOoxml } from '../../src/taskpane/modules/reconciliation/services/table-reconciliation.js';
import {
    buildCommentsPartXml,
    injectCommentIntoParagraphOoxml,
    injectCommentsIntoOoxml,
    injectCommentsIntoPackage
} from '../../src/taskpane/modules/reconciliation/services/comment-engine.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const BASELINE_PATH = path.join(__dirname, '..', 'fixtures', 'phase4-golden-baseline.json');
const LATEST_PATH = path.join(__dirname, '..', 'fixtures', 'phase4-golden-latest.json');
const FIXED_DATE_ISO = '2026-02-07T00:00:00.000Z';

const NativeDate = Date;
const fixedMs = Date.parse(FIXED_DATE_ISO);

class FixedDate extends NativeDate {
    constructor(...args) {
        if (args.length === 0) {
            super(fixedMs);
            return;
        }
        super(...args);
    }

    static now() {
        return fixedMs;
    }
}

function sha256(value) {
    return crypto.createHash('sha256').update(String(value)).digest('hex');
}

function summarizeXml(xml) {
    return {
        length: xml.length,
        hash: sha256(xml)
    };
}

function withFixedDate(fn) {
    globalThis.Date = FixedDate;
    return Promise.resolve()
        .then(fn)
        .finally(() => {
            globalThis.Date = NativeDate;
        });
}

function markdownTable(headers, rows) {
    const lines = [];
    lines.push(`| ${headers.join(' | ')} |`);
    lines.push(`| ${headers.map(() => '---').join(' | ')} |`);
    rows.forEach(row => lines.push(`| ${row.join(' | ')} |`));
    return lines.join('\n');
}

function buildSimplePackage(documentText) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="${NS_W}">
        <w:body>
          <w:p><w:r><w:t>${documentText}</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

async function runCases() {
    const cases = {};

    const formatAddOxml = `<w:document xmlns:w="${NS_W}"><w:body><w:p><w:r><w:t>This is sample text.</w:t></w:r></w:p></w:body></w:document>`;
    resetRevisionIdCounter(1000);
    const formatAddOriginal = ingestOoxml(formatAddOxml).acceptedText;
    const formatAddResult = await applyRedlineToOxml(formatAddOxml, formatAddOriginal, 'This is **sample** text.', {
        author: 'Phase4Golden',
        generateRedlines: true
    });
    cases.format_only_add = {
        hasChanges: formatAddResult.hasChanges,
        hasInsertion: formatAddResult.oxml.includes('w:ins'),
        hasRPrChange: formatAddResult.oxml.includes('w:rPrChange'),
        ...summarizeXml(formatAddResult.oxml)
    };

    const formatRemoveOxml = `<w:document xmlns:w="${NS_W}"><w:body><w:p><w:r><w:t>This is </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>sample</w:t></w:r><w:r><w:t> text.</w:t></w:r></w:p></w:body></w:document>`;
    resetRevisionIdCounter(1000);
    const formatRemoveOriginal = ingestOoxml(formatRemoveOxml).acceptedText;
    const formatRemoveResult = await applyRedlineToOxml(formatRemoveOxml, formatRemoveOriginal, 'This is sample text.', {
        author: 'Phase4Golden',
        generateRedlines: true
    });
    cases.format_only_remove = {
        hasChanges: formatRemoveResult.hasChanges,
        hasRPrChange: formatRemoveResult.oxml.includes('w:rPrChange'),
        ...summarizeXml(formatRemoveResult.oxml)
    };

    const mixedOxml = `<w:document xmlns:w="${NS_W}"><w:body><w:p><w:r><w:t>The quick brown fox jumps.</w:t></w:r></w:p></w:body></w:document>`;
    resetRevisionIdCounter(1000);
    const mixedOriginal = ingestOoxml(mixedOxml).acceptedText;
    const mixedResult = await applyRedlineToOxml(mixedOxml, mixedOriginal, 'The quick red fox hopped.', {
        author: 'Phase4Golden',
        generateRedlines: true
    });
    cases.mixed_insert_delete = {
        hasChanges: mixedResult.hasChanges,
        hasInsertion: mixedResult.oxml.includes('w:ins'),
        hasDeletion: mixedResult.oxml.includes('w:del'),
        ...summarizeXml(mixedResult.oxml)
    };

    const listOxml = `<w:document xmlns:w="${NS_W}"><w:body><w:p><w:r><w:t>List seed</w:t></w:r></w:p></w:body></w:document>`;
    resetRevisionIdCounter(1000);
    const listOriginal = ingestOoxml(listOxml).acceptedText;
    const listResult = await applyRedlineToOxml(
        listOxml,
        listOriginal,
        '- Alpha\n  - Beta\n- Gamma',
        { author: 'Phase4Golden', generateRedlines: true }
    );
    cases.list_generation = {
        hasChanges: listResult.hasChanges,
        hasNumberingRelationship: listResult.oxml.includes('numbering.xml'),
        ...summarizeXml(listResult.oxml)
    };

    const originalTableData = {
        headers: ['Name', 'Qty'],
        rows: [
            ['Apple', '1'],
            ['Berry', '2']
        ],
        hasHeader: true
    };
    const tableXml = generateTableOoxml(originalTableData, { generateRedlines: false, author: 'Phase4Golden' });
    const tableDocOxml = `<w:document xmlns:w="${NS_W}"><w:body>${tableXml}</w:body></w:document>`;
    resetRevisionIdCounter(1000);
    const tableOriginal = ingestOoxml(tableDocOxml).acceptedText;
    const tableResult = await applyRedlineToOxml(
        tableDocOxml,
        tableOriginal,
        '| Name | Qty |\n| --- | --- |\n| Apple | 3 |\n| Citrus | 4 |',
        { author: 'Phase4Golden', generateRedlines: true }
    );
    cases.table_reconcile = {
        hasChanges: tableResult.hasChanges,
        hasTable: tableResult.oxml.includes('<w:tbl'),
        ...summarizeXml(tableResult.oxml)
    };

    const textToTableOxml = `<w:document xmlns:w="${NS_W}"><w:body><w:p><w:r><w:t>Replace this with a table</w:t></w:r></w:p></w:body></w:document>`;
    resetRevisionIdCounter(1000);
    const textToTableOriginal = ingestOoxml(textToTableOxml).acceptedText;
    const textToTableResult = await applyRedlineToOxml(
        textToTableOxml,
        textToTableOriginal,
        '| Col A | Col B |\n| --- | --- |\n| A1 | B1 |\n| A2 | B2 |',
        { author: 'Phase4Golden', generateRedlines: true }
    );
    cases.text_to_table_transform = {
        hasChanges: textToTableResult.hasChanges,
        hasTable: textToTableResult.oxml.includes('<w:tbl') || textToTableResult.oxml.includes('<w:ins'),
        ...summarizeXml(textToTableResult.oxml)
    };

    const commentDocOxml = `<w:document xmlns:w="${NS_W}"><w:body><w:p><w:r><w:t>Paragraph with target_one and target_two.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph target_three.</w:t></w:r></w:p></w:body></w:document>`;
    resetRevisionIdCounter(1000);
    const commentDocResult = injectCommentsIntoOoxml(commentDocOxml, [
        { paragraphIndex: 1, textToFind: 'target_one', commentContent: 'First comment' },
        { paragraphIndex: 2, textToFind: 'target_three', commentContent: 'Second comment' }
    ], { author: 'Phase4Golden' });
    cases.comment_document_injection = {
        commentsApplied: commentDocResult.commentsApplied,
        warningCount: commentDocResult.warnings.length,
        ...summarizeXml(commentDocResult.oxml),
        commentsXmlHash: sha256(commentDocResult.commentsXml || '')
    };

    const paragraphOxml = `<w:p xmlns:w="${NS_W}"><w:r><w:t>Paragraph level target marker.</w:t></w:r></w:p>`;
    resetRevisionIdCounter(1000);
    const paragraphCommentResult = injectCommentIntoParagraphOoxml(
        paragraphOxml,
        'target marker',
        'Paragraph comment',
        { author: 'Phase4Golden' }
    );
    cases.comment_paragraph_injection = {
        success: !!paragraphCommentResult.success,
        hasPackage: (paragraphCommentResult.package || '').includes('pkg:package'),
        packageHash: sha256(paragraphCommentResult.package || ''),
        packageLength: (paragraphCommentResult.package || '').length
    };

    const commentsXml = buildCommentsPartXml([
        { id: 1, content: 'Packaged comment', author: 'Phase4Golden', date: FIXED_DATE_ISO }
    ]);
    const packageOxml = buildSimplePackage('Package target text');
    resetRevisionIdCounter(1000);
    const packageCommentResult = injectCommentsIntoPackage(packageOxml, commentsXml);
    cases.comment_package_injection = {
        hasCommentsPart: packageCommentResult.includes('/word/comments.xml'),
        hasCommentsRel: packageCommentResult.includes('relationships/comments'),
        ...summarizeXml(packageCommentResult)
    };

    return {
        generatedAt: FIXED_DATE_ISO,
        nodeVersion: process.version,
        cases
    };
}

function compareBaselines(expected, actual) {
    const expectedCases = expected.cases || {};
    const actualCases = actual.cases || {};
    const allCaseNames = new Set([...Object.keys(expectedCases), ...Object.keys(actualCases)]);
    const mismatches = [];

    allCaseNames.forEach(caseName => {
        const expectedCase = expectedCases[caseName];
        const actualCase = actualCases[caseName];
        if (!expectedCase || !actualCase) {
            mismatches.push(`${caseName}: missing case`);
            return;
        }

        const keys = new Set([...Object.keys(expectedCase), ...Object.keys(actualCase)]);
        keys.forEach(key => {
            if (expectedCase[key] !== actualCase[key]) {
                mismatches.push(`${caseName}.${key}: expected=${expectedCase[key]} actual=${actualCase[key]}`);
            }
        });
    });

    return mismatches;
}

async function writeJson(filePath, value) {
    await fs.mkdir(path.dirname(filePath), { recursive: true });
    await fs.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

async function main() {
    const shouldUpdate = process.argv.includes('--update');

    await withFixedDate(async () => {
        const actual = await runCases();
        await writeJson(LATEST_PATH, actual);

        if (shouldUpdate) {
            await writeJson(BASELINE_PATH, actual);
            console.log(`Updated golden baseline: ${BASELINE_PATH}`);
            return;
        }

        const expectedRaw = await fs.readFile(BASELINE_PATH, 'utf8');
        const expected = JSON.parse(expectedRaw);
        const mismatches = compareBaselines(expected, actual);

        assert.equal(mismatches.length, 0, `Golden guardrail mismatches:\n${mismatches.join('\n')}`);
        console.log('PASS: phase4 golden guardrail');
    });
}

main().catch(error => {
    console.error('FAIL:', error.message);
    process.exit(1);
});
