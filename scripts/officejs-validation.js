/* global Office, Word */
// Development-only entry point; never included in ordinary taskpane builds.
import { executePureOoxmlBatch, captureWordSourceBaseline } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { applyRedlineChangesToWordContext } from '../src/taskpane/modules/docx-redline-js-integration/word-redline-runner.js';
import { planAgenticListOperations } from '../src/taskpane/modules/commands/list-operation-plan.js';
import { initAgenticTools, executeInsertListItem } from '../src/taskpane/modules/commands/agentic-tools.js';

const status = message => { document.getElementById('status').textContent = message; };
function assert(condition, message) { if (!condition) throw new Error(message); }
function createTimingSamples() {
    const phases = {};
    return {
        record({ phase, durationMs }) { phases[phase] = durationMs; },
        snapshot() { return { sampleCount: 1, ...phases }; }
    };
}
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
        const timing = createTimingSamples();
        // Count actual calls while forwarding to genuine Word proxies.
        const scope = {
            getOoxml() { reads++; return body.getOoxml(); },
            insertOoxml(...args) { writes++; return body.insertOoxml(...args); }
        };
        const result = await executePureOoxmlBatch(context, scope, operations, {
            author: 'WP6 reviewer', onTiming: timing.record
        });
        return { status: result.status, written: result.written, hasChanges: result.hasChanges,
            writeAttempted: result.writeAttempted, mutationOutcome: result.mutationOutcome,
            errorCode: result.error?.code, reads, writes, timing: timing.snapshot() };
    });
}
function getChangeTrackingMode(name) {
    if (name === 'off') return Word.ChangeTrackingMode.off;
    if (name === 'trackAll') return Word.ChangeTrackingMode.trackAll;
    throw new Error(`Unsupported test change-tracking mode: ${name}`);
}
function describeChangeTrackingMode(mode) {
    if (mode === Word.ChangeTrackingMode.off) return 'off';
    if (mode === Word.ChangeTrackingMode.trackAll) return 'trackAll';
    return String(mode);
}
async function prepareProductionSource(testCase) {
    const preparation = testCase.sourcePreparation || [];
    if (!preparation.length) return [];
    return Word.run(async context => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items/text');
        await context.sync();
        const applied = [];
        for (const step of preparation) {
            const matches = paragraphs.items.filter(paragraph => String(paragraph.text || '').replace(/[\r\n]+$/, '') === step.paragraphText);
            assert(matches.length === 1, `Source preparation target must be unique: ${step.paragraphText}`);
            const paragraph = matches[0];
            if (step.type === 'list-level') {
                paragraph.load('listItem');
                await context.sync();
                assert(paragraph.listItem && !paragraph.listItem.isNullObject, `Expected list item for ${step.paragraphText}`);
                paragraph.listItem.level = step.level;
                await context.sync();
            } else if (step.type === 'list-numbering') {
                paragraph.load('list');
                await context.sync();
                const numbering = Word.ListNumbering[step.numbering];
                assert(numbering, `Unknown Word list numbering style: ${step.numbering}`);
                paragraph.list.setLevelNumbering(step.level, numbering);
                await context.sync();
            } else {
                throw new Error(`Unsupported production fixture preparation: ${step.type}`);
            }
            applied.push({ type: step.type, paragraphText: step.paragraphText, level: step.level, numbering: step.numbering });
        }
        return applied;
    });
}
async function runProductionListInsertion(request, productionOptions = {}) {
    let reads = 0, writes = 0;
    const nativeParagraphInserts = [];
    const nativeListLevelWrites = [];
    const timing = createTimingSamples();
    const nativeRun = Word.run;
    if (productionOptions.priorTrackingMode) {
        const desired = getChangeTrackingMode(productionOptions.priorTrackingMode);
        await nativeRun.call(Word, async context => {
            context.document.changeTrackingMode = desired;
            await context.sync();
        });
    }
    const trackingModeBefore = await nativeRun.call(Word, async context => {
        context.document.load('changeTrackingMode');
        await context.sync();
        return describeChangeTrackingMode(context.document.changeTrackingMode);
    });
    const trackingTransitions = [];
    const instrumentParagraph = paragraph => new Proxy(paragraph, {
        get(target, property) {
            if (property === 'insertParagraph') {
                return (...args) => {
                    nativeParagraphInserts.push({ text: args[0], location: args[1] });
                    return instrumentParagraph(target.insertParagraph(...args));
                };
            }
            if (property === 'listItem') {
                const listItem = Reflect.get(target, property, target);
                return new Proxy(listItem, {
                    set(item, key, value) {
                        if (key === 'level') nativeListLevelWrites.push(value);
                        return Reflect.set(item, key, value, item);
                    }
                });
            }
            const value = Reflect.get(target, property, target);
            return typeof value === 'function' ? value.bind(target) : value;
        },
        set(target, property, value) { return Reflect.set(target, property, value, target); }
    });
    const instrumentParagraphCollection = paragraphs => new Proxy(paragraphs, {
        get(target, property) {
            if (property === 'items') {
                const items = Reflect.get(target, property, target);
                return new Proxy(items, {
                    get(collection, index) {
                        const item = Reflect.get(collection, index, collection);
                        return typeof index === 'string' && /^\d+$/.test(index) ? instrumentParagraph(item) : item;
                    }
                });
            }
            const value = Reflect.get(target, property, target);
            return typeof value === 'function' ? value.bind(target) : value;
        }
    });
    initAgenticTools({
        getRequestSignal: () => null, loadRedlineSetting: () => productionOptions.redlineEnabled !== false,
        loadRedlineAuthor: () => 'WP6 reviewer',
        setChangeTrackingForAi: async (context, redlineEnabled) => {
            const document = context.document;
            document.load('changeTrackingMode');
            await context.sync();
            const originalMode = document.changeTrackingMode;
            const desiredMode = redlineEnabled ? Word.ChangeTrackingMode.trackAll : Word.ChangeTrackingMode.off;
            const changed = originalMode !== desiredMode;
            if (changed) {
                document.changeTrackingMode = desiredMode;
                await context.sync();
            }
            trackingTransitions.push({ phase: 'set', redlineEnabled, from: describeChangeTrackingMode(originalMode), to: describeChangeTrackingMode(desiredMode), changed });
            return { available: true, originalMode, changed };
        },
        restoreChangeTracking: async (context, state) => {
            if (!state?.available || !state.changed || state.originalMode === null) return;
            context.document.changeTrackingMode = state.originalMode;
            await context.sync();
            trackingTransitions.push({ phase: 'restore', to: describeChangeTrackingMode(state.originalMode) });
        }
    });
    try {
        const sourceBaseline = await nativeRun.call(Word, async context => {
            const source = context.document.body.getOoxml();
            await context.sync();
            return captureWordSourceBaseline(source.value);
        });
        // Forward production tool execution to genuine Word proxies while counting transport.
        Word.run = callback => nativeRun.call(Word, async context => {
            const documentProxy = context.document;
            const body = documentProxy.body;
            let adapterStartedAt = null;
            let sourceSyncPending = false;
            let preparationStartedAt = null;
            let insertStartedAt = null;
            let insertSyncPending = false;
            const now = () => performance.now();
            const recordAdapterSync = async () => {
                try {
                    return await context.sync();
                } finally {
                    const finishedAt = now();
                    if (sourceSyncPending) {
                        timing.record({ phase: 'sourceReadSync', durationMs: finishedAt - adapterStartedAt });
                        preparationStartedAt = finishedAt;
                        sourceSyncPending = false;
                    }
                    if (insertSyncPending) {
                        timing.record({ phase: 'portablePreparation', durationMs: Math.max(0, insertStartedAt - preparationStartedAt) });
                        timing.record({ phase: 'insertSync', durationMs: finishedAt - insertStartedAt });
                        timing.record({ phase: 'adapterTotal', durationMs: finishedAt - adapterStartedAt });
                        insertSyncPending = false;
                    }
                }
            };
            return callback({ sync: recordAdapterSync, document: {
                load: properties => documentProxy.load(properties),
                get changeTrackingMode() { return documentProxy.changeTrackingMode; },
                set changeTrackingMode(value) { documentProxy.changeTrackingMode = value; },
                body: { paragraphs: instrumentParagraphCollection(body.paragraphs),
                    getOoxml() {
                        reads++;
                        adapterStartedAt = now();
                        sourceSyncPending = true;
                        return body.getOoxml();
                    },
                    insertOoxml(...args) {
                        writes++;
                        const startedAt = adapterStartedAt !== null && preparationStartedAt !== null ? now() : null;
                        const result = body.insertOoxml(...args);
                        if (startedAt !== null) {
                            insertStartedAt = startedAt;
                            insertSyncPending = true;
                        }
                        return result;
                    }
                }
            } });
        });
        const toolStartedAt = performance.now();
        const result = await executeInsertListItem(request.afterParagraphIndex, request.text, request.indentLevel, sourceBaseline);
        timing.record({ phase: 'agenticToolTotal', durationMs: performance.now() - toolStartedAt });
        const trackingModeAfter = await nativeRun.call(Word, async context => {
            context.document.load('changeTrackingMode');
            await context.sync();
            return describeChangeTrackingMode(context.document.changeTrackingMode);
        });
        const route = nativeParagraphInserts.length ? 'native' : writes ? 'canonical' : 'none';
        if (productionOptions.expectedRoute) assert(route === productionOptions.expectedRoute,
            `Expected ${productionOptions.expectedRoute} insertion route, observed ${route}`);
        if (productionOptions.expectedParagraphInsertCalls != null) assert(nativeParagraphInserts.length === productionOptions.expectedParagraphInsertCalls,
            `Expected ${productionOptions.expectedParagraphInsertCalls} native paragraph inserts, got ${nativeParagraphInserts.length}`);
        if (productionOptions.expectedBodyInsertCalls != null) assert(writes === productionOptions.expectedBodyInsertCalls,
            `Expected ${productionOptions.expectedBodyInsertCalls} OOXML body inserts, got ${writes}`);
        if (productionOptions.expectedListLevel != null) assert(nativeListLevelWrites.at(-1) === productionOptions.expectedListLevel,
            `Expected native list level ${productionOptions.expectedListLevel}, got ${nativeListLevelWrites.at(-1)}`);
        if (productionOptions.expectedTrackingModeAfter) assert(trackingModeAfter === productionOptions.expectedTrackingModeAfter,
            `Expected restored change-tracking mode ${productionOptions.expectedTrackingModeAfter}, got ${trackingModeAfter}`);
        return { status: result.status, written: result.written, writeAttempted: result.writeAttempted,
            mutationOutcome: result.mutationOutcome, errorCode: result.error?.code, reads, writes,
            route, nativeParagraphInserts, nativeListLevelWrites, trackingModeBefore, trackingModeAfter, trackingTransitions,
            timing: timing.snapshot() };
    } finally { Word.run = nativeRun; }
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
            const sourcePreparation = await prepareProductionSource(testCase);
            if (testCase.usesPreparedSource) {
                assert(sourcePreparation.length > 0, `${testCase.name}: prepared source was declared but no preparation ran`);
                await post(`/prepared-source/${testCase.name}`, await exportDocx(), 'application/octet-stream');
            }
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
            const applied = testCase.agenticRequest?.tool === 'insert_list_item'
                ? await runProductionListInsertion(testCase.agenticRequest, testCase.productionInsert)
                : await runBatch(testCase.agenticRequest
                    ? source => planAgenticListOperations(source, testCase.agenticRequest)
                    : testCase.nativeOperations);
            if (testCase.productionInsert?.expectedRoute === 'native') {
                assert(applied.status === 'ok' && applied.written && applied.reads === 1,
                    `${testCase.name}: expected one successful source read and native mutation: ${JSON.stringify(applied)}`);
                assert(applied.route === 'native' && applied.writes === 0 && applied.nativeParagraphInserts.length === 1,
                    `${testCase.name}: expected one native paragraph insert and no OOXML body write`);
                checks.push({ case: testCase.name, view: 'actual-officejs-native-insertion', ...applied });
            } else {
                assert(applied.status === 'ok' && applied.written && applied.reads === 1 && applied.writes === 1,
                    `${testCase.name}: expected one successful read/write: ${JSON.stringify(applied)}`);
                checks.push({ case: testCase.name, view: 'actual-officejs-insertion', ...applied });
            }
            const noOp = await runBatch([]);
            assert(!noOp.written && noOp.reads === 1 && noOp.writes === 0, 'No-op inserted into Word');
            checks.push({ case: testCase.name, view: 'no-op', ...noOp });
            const refused = await runBatch([{ type: 'replace', target: { paragraphIndex: 999999 }, modified: 'must not appear' }]);
            assert(refused.status === 'error' && !refused.written && refused.reads === 1 && refused.writes === 0,
                'Failed preparation inserted into Word');
            checks.push({ case: testCase.name, view: 'preparation-refusal', ...refused });
            if (testCase.agenticRequest) {
                const mixed = await runBatch(source => [
                    { type: 'replace', target: { index: 1, exactText: source.paragraphs[0].exactText }, modified: 'Must roll back.' },
                    { type: 'replace', target: { index: 999999 }, modified: 'Must not appear.' }
                ]);
                assert(mixed.status === 'error' && !mixed.written && !mixed.writeAttempted && mixed.writes === 0,
                    'Mixed valid/invalid list-host batch inserted into Word');
                checks.push({ case: testCase.name, view: 'mixed-batch-refusal', ...mixed });
            }
            await post(`/artifact/${testCase.name}`, await exportDocx(), 'application/octet-stream');
        }
        await post('/result', JSON.stringify({ status: 'passed', host: info, diagnostics: Office.context.diagnostics, checks }));
        status('Office.js transport checks passed. Exported documents require independent Word verification.');
    } catch (error) {
        status(`Failed: ${error.message}`);
        await post('/result', JSON.stringify({ status: 'failed', error: error.message })).catch(() => {});
    }
});
