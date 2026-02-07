import '../setup-xml-provider.mjs';
import assert from 'assert';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

import { NS_W, resetRevisionIdCounter } from '../../src/taskpane/modules/reconciliation/core/types.js';
import { applyRedlineToOxml } from '../../src/taskpane/modules/reconciliation/engine/oxml-engine.js';
import { ingestOoxml } from '../../src/taskpane/modules/reconciliation/pipeline/ingestion.js';
import { generateTableOoxml } from '../../src/taskpane/modules/reconciliation/services/table-reconciliation.js';
import { injectCommentsIntoOoxml } from '../../src/taskpane/modules/reconciliation/services/comment-engine.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const BASELINE_PATH = path.join(__dirname, '..', 'fixtures', 'phase4-perf-baseline.json');
const LATEST_PATH = path.join(__dirname, '..', 'fixtures', 'phase4-perf-latest.json');
const FIXED_DATE_ISO = '2026-02-07T00:00:00.000Z';
const DEFAULT_THRESHOLD = 0.05;

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

function withFixedDate(fn) {
    globalThis.Date = FixedDate;
    return Promise.resolve()
        .then(fn)
        .finally(() => {
            globalThis.Date = NativeDate;
        });
}

function average(values) {
    if (values.length === 0) return 0;
    return values.reduce((acc, value) => acc + value, 0) / values.length;
}

function median(values) {
    if (values.length === 0) return 0;
    const sorted = [...values].sort((a, b) => a - b);
    const middle = Math.floor(sorted.length / 2);
    if (sorted.length % 2 === 0) {
        return (sorted[middle - 1] + sorted[middle]) / 2;
    }
    return sorted[middle];
}

function elapsedMs(startNanos, endNanos) {
    return Number(endNanos - startNanos) / 1_000_000;
}

function buildLongParagraphDocument(runCount = 240) {
    const runTexts = Array.from({ length: runCount }, (_, index) => `segment_${index}_token `);
    const runsXml = runTexts.map(text => `<w:r><w:t>${text}</w:t></w:r>`).join('');
    const oxml = `<w:document xmlns:w="${NS_W}"><w:body><w:p>${runsXml}</w:p></w:body></w:document>`;
    return { oxml, text: runTexts.join('') };
}

function buildLargeTableData(rows = 40, cols = 10) {
    const headers = Array.from({ length: cols }, (_, col) => `H${col + 1}`);
    const rowData = Array.from({ length: rows }, (_, row) =>
        Array.from({ length: cols }, (_, col) => `R${row + 1}C${col + 1}`)
    );
    return { headers, rows: rowData, hasHeader: true };
}

function toMarkdownTable(headers, rows) {
    const lines = [];
    lines.push(`| ${headers.join(' | ')} |`);
    lines.push(`| ${headers.map(() => '---').join(' | ')} |`);
    rows.forEach(row => lines.push(`| ${row.join(' | ')} |`));
    return lines.join('\n');
}

function buildCommentDocument(paragraphCount = 30) {
    const paragraphs = Array.from({ length: paragraphCount }, (_, index) => {
        const number = index + 1;
        return `<w:p><w:r><w:t>Paragraph ${number} has target_${number} marker for comment insertion.</w:t></w:r></w:p>`;
    }).join('');
    return `<w:document xmlns:w="${NS_W}"><w:body>${paragraphs}</w:body></w:document>`;
}

async function benchmarkScenario(name, iterations, scenarioFn) {
    await scenarioFn();
    const samples = [];

    for (let i = 0; i < iterations; i++) {
        const start = process.hrtime.bigint();
        await scenarioFn();
        const end = process.hrtime.bigint();
        samples.push(elapsedMs(start, end));
    }

    return {
        name,
        iterations,
        minMs: Number(Math.min(...samples).toFixed(3)),
        medianMs: Number(median(samples).toFixed(3)),
        avgMs: Number(average(samples).toFixed(3)),
        maxMs: Number(Math.max(...samples).toFixed(3))
    };
}

async function runBenchmarks(iterations = 4) {
    const { oxml: longParagraphOxml, text: longParagraphText } = buildLongParagraphDocument();
    const modifiedLongText = longParagraphText.replace('segment_15_token', 'segment_15_token_UPDATED') + ' trailing addition';

    const longParagraphScenario = async () => {
        resetRevisionIdCounter(1000);
        const result = await applyRedlineToOxml(longParagraphOxml, longParagraphText, modifiedLongText, {
            author: 'Phase4Perf',
            generateRedlines: true
        });
        assert.equal(typeof result.hasChanges, 'boolean');
    };

    const originalTable = buildLargeTableData();
    const updatedRows = originalTable.rows.map((row, rowIndex) =>
        row.map((cell, colIndex) => {
            if ((rowIndex + colIndex) % 9 === 0) return `${cell}_UPD`;
            return cell;
        })
    );
    updatedRows.push(Array.from({ length: originalTable.headers.length }, (_, col) => `NEW_${col + 1}`));

    const tableOxml = generateTableOoxml(originalTable, { generateRedlines: false, author: 'Phase4Perf' });
    const tableDocumentOxml = `<w:document xmlns:w="${NS_W}"><w:body>${tableOxml}</w:body></w:document>`;
    const tableOriginalText = ingestOoxml(tableDocumentOxml).acceptedText;
    const tableMarkdown = toMarkdownTable(originalTable.headers, updatedRows);

    const tableScenario = async () => {
        resetRevisionIdCounter(1000);
        const result = await applyRedlineToOxml(tableDocumentOxml, tableOriginalText, tableMarkdown, {
            author: 'Phase4Perf',
            generateRedlines: true
        });
        assert.equal(typeof result.hasChanges, 'boolean');
    };

    const commentOxml = buildCommentDocument(35);
    const comments = Array.from({ length: 25 }, (_, index) => ({
        paragraphIndex: index + 1,
        textToFind: `target_${index + 1}`,
        commentContent: `Comment ${index + 1}`
    }));

    const commentScenario = async () => {
        resetRevisionIdCounter(1000);
        const result = injectCommentsIntoOoxml(commentOxml, comments, { author: 'Phase4Perf' });
        assert.equal(result.commentsApplied, 25);
    };

    return [
        await benchmarkScenario('long_paragraph_redline', iterations, longParagraphScenario),
        await benchmarkScenario('table_reconciliation', iterations, tableScenario),
        await benchmarkScenario('multi_comment_injection', iterations, commentScenario)
    ];
}

async function writeJson(filePath, value) {
    await fs.mkdir(path.dirname(filePath), { recursive: true });
    await fs.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
}

function comparePerfBudget(baseline, current, threshold) {
    const baselineMap = new Map((baseline.results || []).map(result => [result.name, result]));
    const regressions = [];

    (current.results || []).forEach(result => {
        const baselineResult = baselineMap.get(result.name);
        if (!baselineResult) {
            regressions.push(`${result.name}: missing baseline`);
            return;
        }

        const allowed = baselineResult.medianMs * (1 + threshold);
        if (result.medianMs > allowed) {
            regressions.push(
                `${result.name}: median ${result.medianMs}ms > allowed ${allowed.toFixed(3)}ms (baseline ${baselineResult.medianMs}ms)`
            );
        }
    });

    return regressions;
}

async function main() {
    const shouldSaveBaseline = process.argv.includes('--save-baseline');
    const shouldVerify = process.argv.includes('--verify');
    const thresholdArg = process.argv.find(arg => arg.startsWith('--threshold='));
    const threshold = thresholdArg ? Number(thresholdArg.split('=')[1]) : DEFAULT_THRESHOLD;

    await withFixedDate(async () => {
        const results = await runBenchmarks();
        const payload = {
            generatedAt: FIXED_DATE_ISO,
            nodeVersion: process.version,
            iterations: results[0]?.iterations || 0,
            threshold,
            results
        };

        await writeJson(LATEST_PATH, payload);

        if (shouldSaveBaseline) {
            await writeJson(BASELINE_PATH, payload);
            console.log(`Saved perf baseline: ${BASELINE_PATH}`);
            return;
        }

        if (shouldVerify) {
            const baselineRaw = await fs.readFile(BASELINE_PATH, 'utf8');
            const baseline = JSON.parse(baselineRaw);
            const regressions = comparePerfBudget(baseline, payload, threshold);
            assert.equal(regressions.length, 0, `Performance regressions:\n${regressions.join('\n')}`);
            console.log(`PASS: phase4 perf harness (threshold ${(threshold * 100).toFixed(1)}%)`);
            return;
        }

        console.log('Perf results (no baseline assertion):');
        payload.results.forEach(result => {
            console.log(`- ${result.name}: median=${result.medianMs}ms avg=${result.avgMs}ms min=${result.minMs}ms max=${result.maxMs}ms`);
        });
    });
}

main().catch(error => {
    console.error('FAIL:', error.message);
    process.exit(1);
});
