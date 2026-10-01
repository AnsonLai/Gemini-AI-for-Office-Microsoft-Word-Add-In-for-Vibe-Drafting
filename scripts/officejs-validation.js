/* global Office, Word */
// Development-only entry point; never included in ordinary taskpane builds.
import { executePureOoxmlBatch, captureWordSourceBaseline } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { applyRedlineChangesToWordContext } from '../src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js';
import { planAgenticListOperations } from '../src/taskpane/modules/commands/list-operation-plan.js';

const status = message => { document.getElementById('status').textContent = message; };
function assert(condition, message) { if (!condition) throw new Error(message); }
function officeCall(action) {
    return new Promise((resolve, reject) => action(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error(result.error?.code || 'OFFICE_ASYNC_FAILED'));
    }));
}
async function exportDocx() {
    const file = await officeCall(cb => Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, cb));
    try {
        const slices = [];
        for (let i = 0; i < file.sliceCount; i++) {
            slices.push(new Uint8Array((await officeCall(cb => file.getSliceAsync(i, cb))).data));
        }
        const bytes = new Uint8Array(slices.reduce((n, s) => n + s.length, 0));
        let offset = 0;
        for (const slice of slices) { bytes.set(slice, offset); offset += slice.length; }
        return bytes;
    } finally { await officeCall(cb => file.closeAsync(cb)); }
}
async function post(route, body, type = 'application/json') {
    const response = await fetch(route, { method: 'POST', headers: { 'Content-Type': type }, body });
    if (!response.ok) throw new Error(`Local evidence server HTTP ${response.status}`);
}
async function runBatch(operations) {
    return Word.run(async context => {
        const body = context.document.body;
        let reads = 0, writes = 0;
        // Count actual calls while forwarding to genuine Word proxies.
        const scope = {
            getOoxml() { reads++; return body.getOoxml(); },
            insertOoxml(...args) { writes++; return body.insertOoxml(...args); }
        };
        const result = await executePureOoxmlBatch(context, scope, operations, { author: 'WP6 reviewer' });
        return { status: result.status, written: result.written, hasChanges: result.hasChanges,
            writeAttempted: result.writeAttempted, mutationOutcome: result.mutationOutcome,
            errorCode: result.error?.code, reads, writes };
    });
}
Office.onReady(async info => {
    if (info.host !== Office.HostType.Word) { status('This lane requires desktop Word.'); return; }
    try {
        const response = await fetch('/claim', { method: 'POST' });
        if (response.status === 409) { status('This run has already been claimed; no edits performed.'); return; }
        assert(response.ok, 'Could not claim local run');
        const cases = await response.json();
        const checks = [];
        for (const testCase of cases) {
            status(`Running ${testCase.name}`);
            const fixture = await (await fetch(`/source/${testCase.name}`)).text();
            // Seed only this disposable test document; not a production edit path.
            await Word.run(async context => {
                context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
                context.document.body.insertFileFromBase64(fixture, Word.InsertLocation.replace);
                await context.sync();
            });
            if (testCase.name === 'word-addin-plain-replacement') {
                const stale = await Word.run(async context => {
                    const body = context.document.body;
                    const initialOoxml = body.getOoxml();
                    await context.sync();
                    const sourceBaseline = captureWordSourceBaseline(initialOoxml.value);
                    // Simulate an intervening user edit in this disposable fixture.
                    body.insertText('An intervening edit.', Word.InsertLocation.replace);
                    await context.sync();
                    let writes = 0;
                    const result = await applyRedlineChangesToWordContext(context,
                        [{ operation: 'edit_paragraph', paragraphIndex: 1, newContent: 'Must not overwrite.' }],
                        { sourceBaseline, batchRunner: (ctx, scope, operations, options) => executePureOoxmlBatch(ctx, {
                            getOoxml: () => scope.getOoxml(),
                            insertOoxml(...args) { writes++; return scope.insertOoxml(...args); }
                        }, operations, options), onWarn() {} });
                    body.load('text');
                    await context.sync();
                    assert(result.error?.code === 'STALE_DOCUMENT_CONTEXT' && !result.writeAttempted && writes === 0
                        && body.text.replace(/[\r\n]+$/, '') === 'An intervening edit.', 'Stale context overwrote the intervening edit');
                    body.insertFileFromBase64(fixture, Word.InsertLocation.replace);
                    await context.sync();
                    return { errorCode: result.error.code, writes, written: result.written, writeAttempted: result.writeAttempted };
                });
                checks.push({ case: testCase.name, view: 'stale-context-refusal', ...stale });
                const unchanged = await runBatch([{ type: 'replace', target: { exactText: 'Alpha beta gamma.' }, modified: 'Alpha beta gamma.' }]);
                assert(unchanged.status === 'ok' && !unchanged.hasChanges && !unchanged.written && unchanged.reads === 1 && unchanged.writes === 0,
                    `Unchanged paragraph was not a successful no-op: ${JSON.stringify(unchanged)}`);
                checks.push({ case: testCase.name, view: 'unchanged-paragraph', ...unchanged });
            }
            const applied = await runBatch(testCase.agenticRequest
                ? source => planAgenticListOperations(source, testCase.agenticRequest)
                : testCase.nativeOperations);
            assert(applied.status === 'ok' && applied.written && applied.reads === 1 && applied.writes === 1,
                `${testCase.name}: expected one successful read/write`);
            checks.push({ case: testCase.name, view: 'actual-officejs-insertion', ...applied });
            const noOp = await runBatch([]);
            assert(!noOp.written && noOp.reads === 1 && noOp.writes === 0, 'No-op inserted into Word');
            checks.push({ case: testCase.name, view: 'no-op', ...noOp });
            const refused = await runBatch([{ type: 'replace', target: { paragraphIndex: 999999 }, modified: 'must not appear' }]);
            assert(refused.status === 'error' && !refused.written && refused.reads === 1 && refused.writes === 0,
                'Failed preparation inserted into Word');
            checks.push({ case: testCase.name, view: 'preparation-refusal', ...refused });
            await post(`/artifact/${testCase.name}`, await exportDocx(), 'application/octet-stream');
        }
        await post('/result', JSON.stringify({ status: 'passed', host: info, diagnostics: Office.context.diagnostics, checks }));
        status('Office.js transport checks passed. Exported documents require independent Word verification.');
    } catch (error) {
        status(`Failed: ${error.message}`);
        await post('/result', JSON.stringify({ status: 'failed', error: error.message })).catch(() => {});
    }
});
