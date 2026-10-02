/* global Office, Word */
// Development-only transport collector. Mutations go through the same typed
// document tools and source-bound handles used by the task pane.
import * as docxLibrary from '@ansonlai/docx-redline-js';
import * as wordOperationRunner from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { computeWordSourceRevisionToken, detectExportTailIndex } from '../src/taskpane/modules/docx-redline-js-integration/word-source-identity.js';
import { createDocumentInspectionSession } from '../src/taskpane/modules/chat/document-inspection-session.js';
import { createDocumentTools } from '../src/taskpane/modules/chat/document-tools.js';
import { createWordDocumentSourceReader } from '../src/taskpane/modules/chat/word-document-source.js';

const status = message => { document.getElementById('status').textContent = message; };
function assert(condition, message) { if (!condition) throw new Error(message); }
function describeTrackingMode(mode) {
    if (mode === Word.ChangeTrackingMode.off) return 'off';
    if (mode === Word.ChangeTrackingMode.trackAll) return 'trackAll';
    return String(mode);
}
function trackingMode(name) {
    if (name === 'off') return Word.ChangeTrackingMode.off;
    if (name === 'trackAll') return Word.ChangeTrackingMode.trackAll;
    throw new Error(`Unsupported fixture tracking mode: ${name}`);
}
function createTransportCounters() {
    return {
        sourceReads: 0,
        contentWrites: 0,
        nativeBodyTextWrites: 0,
        nativeParagraphInsertCalls: [],
        nativeListLevelWrites: [],
        nativeListMethodCalls: [],
        trackingTransitions: [],
        inspectionSourceSnapshots: 0,
    };
}
function instrumentedRun(stats) {
    return callback => Word.run(async nativeContext => {
        const nativeDocument = nativeContext.document;
        const nativeBody = nativeDocument.body;
        const instrumentListItem = item => item && new Proxy(item, {
            get(raw, property) {
                const value = Reflect.get(raw, property, raw);
                return typeof value === 'function' ? value.bind(raw) : value;
            },
            set(raw, property, value) {
                if (property === 'level') stats.nativeListLevelWrites.push(value);
                return Reflect.set(raw, property, value, raw);
            }
        });
        const instrumentList = list => list && new Proxy(list, {
            get(raw, property) {
                const value = Reflect.get(raw, property, raw);
                if (['setLevelNumbering', 'startNewList', 'attachToList'].includes(String(property)) && typeof value === 'function') {
                    return (...args) => {
                        stats.nativeListMethodCalls.push(String(property));
                        return value.apply(raw, args);
                    };
                }
                return typeof value === 'function' ? value.bind(raw) : value;
            }
        });
        const instrumentParagraph = paragraph => new Proxy(paragraph, {
            get(raw, property) {
                if (property === 'insertParagraph') return (...args) => {
                    stats.nativeParagraphInsertCalls.push({ text: args[0], location: args[1] });
                    return instrumentParagraph(raw.insertParagraph(...args));
                };
                if (property === 'listItem') return instrumentListItem(Reflect.get(raw, property, raw));
                if (property === 'list') return instrumentList(Reflect.get(raw, property, raw));
                const value = Reflect.get(raw, property, raw);
                return typeof value === 'function' ? value.bind(raw) : value;
            },
            set(raw, property, value) { return Reflect.set(raw, property, value, raw); }
        });
        const nativeParagraphs = nativeBody.paragraphs;
        const paragraphCollection = new Proxy(nativeParagraphs, {
            get(raw, property) {
                if (property === 'items') {
                    const items = Reflect.get(raw, property, raw);
                    return new Proxy(items, {
                        get(collection, index) {
                            const item = Reflect.get(collection, index, collection);
                            return typeof index === 'string' && /^\d+$/.test(index) ? instrumentParagraph(item) : item;
                        }
                    });
                }
                const value = Reflect.get(raw, property, raw);
                return typeof value === 'function' ? value.bind(raw) : value;
            }
        });
        const body = new Proxy(nativeBody, {
            get(raw, property) {
                if (property === 'paragraphs') return paragraphCollection;
                if (property === 'getOoxml') return (...args) => {
                    stats.sourceReads++;
                    return raw.getOoxml(...args);
                };
                if (property === 'insertOoxml') return (...args) => {
                    stats.contentWrites++;
                    return raw.insertOoxml(...args);
                };
                if (property === 'insertText' || property === 'insertFileFromBase64') return (...args) => {
                    stats.nativeBodyTextWrites++;
                    return raw[property](...args);
                };
                const value = Reflect.get(raw, property, raw);
                return typeof value === 'function' ? value.bind(raw) : value;
            }
        });
        const nativeDocumentProxy = new Proxy(nativeDocument, {
            get(raw, property) {
                if (property === 'body') return body;
                const value = Reflect.get(raw, property, raw);
                return typeof value === 'function' ? value.bind(raw) : value;
            },
            set(raw, property, value) {
                if (property === 'changeTrackingMode') {
                    stats.trackingTransitions.push({
                        from: describeTrackingMode(Reflect.get(raw, property, raw)),
                        to: describeTrackingMode(value)
                    });
                }
                return Reflect.set(raw, property, value, raw);
            }
        });
        const context = new Proxy(nativeContext, {
            get(raw, property) {
                if (property === 'document') return nativeDocumentProxy;
                const value = Reflect.get(raw, property, raw);
                return typeof value === 'function' ? value.bind(raw) : value;
            }
        });
        return callback(context);
    });
}

async function getTrackingMode(run = callback => Word.run(callback)) {
    return run(async context => {
        context.document.load('changeTrackingMode');
        await context.sync();
        return describeTrackingMode(context.document.changeTrackingMode);
    });
}
async function setTrackingMode(mode) {
    await Word.run(async context => {
        context.document.changeTrackingMode = trackingMode(mode);
        await context.sync();
    });
}
async function exportDocx() {
    const officeCall = action => new Promise((resolve, reject) => action(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error(result.error?.code || 'OFFICE_ASYNC_FAILED'));
    }));
    const file = await officeCall(cb => Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, cb));
    try {
        const slices = [];
        for (let index = 0; index < file.sliceCount; index++) {
            slices.push(new Uint8Array((await officeCall(cb => file.getSliceAsync(index, cb))).data));
        }
        const bytes = new Uint8Array(slices.reduce((length, slice) => length + slice.length, 0));
        let offset = 0;
        for (const slice of slices) { bytes.set(slice, offset); offset += slice.length; }
        return bytes;
    } finally { await officeCall(cb => file.closeAsync(cb)); }
}
async function post(route, body, type = 'application/json') {
    const response = await fetch(route, { method: 'POST', headers: { 'Content-Type': type }, body });
    if (!response.ok) throw new Error(`Local evidence server HTTP ${response.status}`);
}

async function createDirectDocumentTools(testCase, stats) {
    const run = instrumentedRun(stats);
    const documentId = `${testCase.name}:${Office.context.document.url || 'disposable-word-document'}`;
    const readWordSource = createWordDocumentSourceReader({
        loadSupport: async () => [wordOperationRunner, docxLibrary, { createDocumentInspectionSession }],
        getDocumentId: () => documentId,
        run,
    });
    // Cache the one initial capture so the session's first inspection uses this
    // exact Word source, including its scalar host paragraph count.
    let initialSource = await readWordSource();
    const capturedSource = initialSource;
    const readSource = async () => {
        stats.inspectionSourceSnapshots++;
        if (initialSource) {
            const source = initialSource;
            initialSource = null;
            return source;
        }
        return readWordSource();
    };
    const session = createDocumentInspectionSession({
        readSource,
        maxTextChars: 4000,
        computeRevisionToken: computeWordSourceRevisionToken,
        getExcludedParagraphIndex: detectExportTailIndex,
    });
    let checkpointCount = 0;
    const receipts = [];
    const tools = createDocumentTools({
        session,
        getDocumentId: () => documentId,
        getAuthor: () => 'WP4 reviewer',
        getGenerateRedlines: () => testCase.productionOptions?.generateRedlines !== false,
        async persistCheckpoint(sourceOoxml) {
            checkpointCount++;
            await post(`/checkpoint/${testCase.name}`, sourceOoxml, 'application/xml');
            return checkpointCount;
        },
        run,
        onReceipt: receipt => receipts.push(receipt),
    });
    return { tools, session, capturedSource, stats, receipts, get checkpointCount() { return checkpointCount; } };
}

function requestedEdits(testCase) {
    if (testCase.agenticRequest?.type === 'insert_list_item') {
        assert(Number.isInteger(testCase.agenticRequest.afterParagraphIndex), 'List oracle metadata requires the observed Word index.');
        return [{ type: 'insert_list_item', targetText: testCase.agenticRequest.targetText, text: testCase.agenticRequest.text }];
    }
    assert(Array.isArray(testCase.edits) && testCase.edits.length > 0, `${testCase.name}: direct edits are missing.`);
    return testCase.edits;
}
async function handleFor(tools, targetText) {
    const inspected = await tools.inspectDocument({ search: targetText, limit: 8 });
    assert(inspected.status === 'ok', `Inspection failed for ${targetText}: ${JSON.stringify(inspected.error)}`);
    const matches = inspected.targets.filter(target => target.text === targetText && !target.textTruncated);
    assert(matches.length === 1, `Expected one fully inspected target for ${targetText}, found ${matches.length}.`);
    return matches[0].handle;
}
async function applyRequestedEdits(tools, testCase) {
    const edits = [];
    const handles = new Map();
    for (const proposal of requestedEdits(testCase)) {
        let handle = handles.get(proposal.targetText);
        if (!handle) {
            handle = await handleFor(tools, proposal.targetText);
            handles.set(proposal.targetText, handle);
        }
        const { targetText, ...edit } = proposal;
        edits.push({ ...edit, target: handle });
    }
    return tools.applyDocumentEdits({ edits });
}
function assertCommitted(receipt, testCase, stats, checkpointCount) {
    assert(receipt?.success === true && receipt.verified === true && receipt.completion === true,
        `${testCase.name}: document tools did not return a confirmed receipt: ${JSON.stringify(receipt)}`);
    assert(receipt.accountedCount === requestedEdits(testCase).length,
        `${testCase.name}: the receipt did not account for every requested edit: ${JSON.stringify(receipt)}`);
    assert(receipt.written === true && receipt.writeAttempted === true,
        `${testCase.name}: expected a single confirmed body write: ${JSON.stringify(receipt)}`);
    assert(Number.isInteger(receipt.checkpointIndex) && receipt.checkpointIndex > 0 && checkpointCount === 1,
        `${testCase.name}: the same-source recovery checkpoint was not persisted.`);
    assert(receipt.counters?.contentWrites === 1 && receipt.counters?.verificationReads === 1
        && receipt.counters?.sourceReads >= 2,
    `${testCase.name}: unexpected direct-tool receipt counters: ${JSON.stringify(receipt.counters)}`);
    assert(stats.contentWrites === 1, `${testCase.name}: expected one OOXML body write, got ${stats.contentWrites}.`);
    assert(stats.nativeBodyTextWrites === 0 && stats.nativeParagraphInsertCalls.length === 0
        && stats.nativeListLevelWrites.length === 0 && stats.nativeListMethodCalls.length === 0,
    `${testCase.name}: a native edit route was invoked.`);
}
function tableStructureFromSource(source, tableFixture) {
    assert(Number.isInteger(source.hostParagraphCount), 'Word source did not capture the body paragraph membership count.');
    const excludedTailIndex = detectExportTailIndex(source.parts, { hostParagraphCount: source.hostParagraphCount });
    const paragraphs = source.inspection.paragraphs.filter(paragraph => paragraph.inTable !== true
        && paragraph.index !== excludedTailIndex).map(paragraph => paragraph.exactText);
    const state = cells => ({
        paragraphs,
        tables: [{ rows: tableFixture.rows, columns: tableFixture.columns, cells }]
    });
    return {
        source: state(tableFixture.source),
        accepted: state(tableFixture.accepted),
        rejected: state(tableFixture.rejected),
    };
}

async function seedFixture(fixture) {
    await Word.run(async context => {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        context.document.body.insertFileFromBase64(fixture, Word.InsertLocation.replace);
        await context.sync();
    });
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
            } else throw new Error(`Unsupported production fixture preparation: ${step.type}`);
            applied.push({ type: step.type, paragraphText: step.paragraphText, level: step.level, numbering: step.numbering });
        }
        return applied;
    });
}
async function readInterveningEdit() {
    await Word.run(async context => {
        context.document.body.insertText('Intervening user edit. ', Word.InsertLocation.start);
        await context.sync();
    });
}
async function runStaleHandleProbe(testCase, fixture, targetText) {
    const stats = createTransportCounters();
    const harness = await createDirectDocumentTools(testCase, stats);
    const handle = await handleFor(harness.tools, targetText);
    await readInterveningEdit();
    const receipt = await harness.tools.applyDocumentEdits({ edits: [{ type: 'rewrite_paragraph', target: handle, text: targetText }] });
    assert(receipt.success === false && receipt.written === false && receipt.writeAttempted === false,
        `Stale handle wrote to Word: ${JSON.stringify(receipt)}`);
    assert(harness.checkpointCount === 0 && stats.contentWrites === 0,
        'Stale handle must refuse before checkpoint persistence and OOXML insertion.');
    await seedFixture(fixture);
    return { status: receipt.status, errorCode: receipt.error?.code, written: receipt.written, writeAttempted: receipt.writeAttempted, checkpointCount: harness.checkpointCount };
}
async function runNoopAndRefusalControls(testCase, harness) {
    const targetText = harness.capturedSource.inspection.paragraphs.find(paragraph => paragraph.inTable !== true && paragraph.exactText)?.exactText;
    assert(targetText, 'No-op control needs a non-empty source paragraph.');
    const target = await handleFor(harness.tools, targetText);
    const noop = await harness.tools.applyDocumentEdits({ edits: [{ type: 'rewrite_paragraph', target, text: targetText }] });
    assert(noop.success === true && noop.verified === true && noop.written === false && noop.writeAttempted === false,
        `Direct-tool no-op was not verified: ${JSON.stringify(noop)}`);
    assert(noop.counters?.contentWrites === 0 && harness.checkpointCount === 1,
        'A verified no-op must not create a checkpoint or content write.');
    const refusal = await harness.tools.applyDocumentEdits({ edits: [{ type: 'rewrite_paragraph', target: 'expired-handle', text: 'Must not appear.' }] });
    assert(refusal.success === false && refusal.written === false && refusal.writeAttempted === false
        && refusal.counters?.contentWrites === 0 && harness.checkpointCount === 1,
    `Invalid target was not refused before a host write: ${JSON.stringify(refusal)}`);
    return [
        { case: testCase.name, view: 'direct-tool-verified-noop', receipt: noop },
        { case: testCase.name, view: 'direct-tool-invalid-handle-refusal', receipt: refusal },
    ];
}

Office.onReady(async info => {
    if (info.host !== Office.HostType.Word) { status('This lane requires desktop Word.'); return; }
    try {
        docxLibrary.setPlatform(Office.context.platform);
        const response = await fetch('/claim', { method: 'POST' });
        if (response.status === 409) { status('This run has already been claimed; no edits performed.'); return; }
        assert(response.ok, 'Could not claim local run');
        const cases = await response.json();
        const checks = [];
        const fixtureExpectations = {};
        for (const testCase of cases) {
            status(`Running ${testCase.name}`);
            const fixture = await (await fetch(`/source/${testCase.name}`)).text();
            await seedFixture(fixture);
            const sourcePreparation = await prepareProductionSource(testCase);
            if (testCase.usesPreparedSource) {
                assert(sourcePreparation.length > 0, `${testCase.name}: prepared source was declared but no preparation ran`);
                await post(`/prepared-source/${testCase.name}`, await exportDocx(), 'application/octet-stream');
            }

            if (testCase.name === cases[0]?.name) {
                const anchorText = requestedEdits(testCase)[0].targetText;
                const staleProbe = await runStaleHandleProbe(testCase, fixture, anchorText);
                checks.push({ case: testCase.name, view: 'stale-handle-refusal', ...staleProbe });
            }

            await setTrackingMode(testCase.productionOptions?.priorTrackingMode || 'off');
            const trackingModeBefore = await getTrackingMode();
            const stats = createTransportCounters();
            const harness = await createDirectDocumentTools(testCase, stats);
            const source = harness.capturedSource;
            assert(source.inspection?.status === 'ok', `${testCase.name}: canonical source inspection failed.`);
            const receipt = await applyRequestedEdits(harness.tools, testCase);
            const trackingModeAfter = await getTrackingMode();
            assert(trackingModeAfter === trackingModeBefore,
                `${testCase.name}: change-tracking state was not restored (${trackingModeBefore} → ${trackingModeAfter}).`);
            assertCommitted(receipt, testCase, stats, harness.checkpointCount);
            const item = {
                case: testCase.name,
                view: 'actual-officejs-direct-document-tools',
                author: 'WP4 reviewer',
                receipt,
                hostTransport: {
                    sourceReads: stats.sourceReads,
                    inspectionSourceSnapshots: stats.inspectionSourceSnapshots,
                    hostParagraphCount: source.hostParagraphCount,
                    contentWrites: stats.contentWrites,
                    nativeBodyTextWrites: stats.nativeBodyTextWrites,
                    nativeParagraphInsertCalls: stats.nativeParagraphInsertCalls,
                    nativeListLevelWrites: stats.nativeListLevelWrites,
                    nativeListMethodCalls: stats.nativeListMethodCalls,
                    trackingTransitions: stats.trackingTransitions,
                    checkpointCount: harness.checkpointCount
                },
                trackingState: { before: trackingModeBefore, after: trackingModeAfter },
            };
            checks.push(item);
            if (sourcePreparation.length) item.sourcePreparation = sourcePreparation;
            if (testCase.tableStructure) {
                fixtureExpectations[testCase.name] = {
                    expectedBodyStructure: tableStructureFromSource(source, testCase.tableStructure),
                    hostParagraphCount: source.hostParagraphCount,
                    excludedExportTailIndex: detectExportTailIndex(source.parts, { hostParagraphCount: source.hostParagraphCount })
                };
            }
            if (testCase.name === cases[0]?.name) checks.push(...await runNoopAndRefusalControls(testCase, harness));
            await post(`/artifact/${testCase.name}`, await exportDocx(), 'application/octet-stream');
        }
        await post('/result', JSON.stringify({
            status: 'passed', host: info, diagnostics: Office.context.diagnostics, checks, fixtureExpectations
        }));
        status('Direct document-tool transport checks passed. Exported documents require independent Word verification.');
    } catch (error) {
        status(`Failed: ${error.message}`);
        await post('/result', JSON.stringify({ status: 'failed', error: error.message })).catch(() => {});
    }
});
