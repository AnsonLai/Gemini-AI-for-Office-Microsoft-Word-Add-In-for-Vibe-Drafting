/* global Office, Word */
// Development-only transport collector. Mutations go through the same typed
// document tools and source-bound handles used by the task pane.
import * as docxLibrary from '@ansonlai/docx-redline-js';
import { extractDocumentPartsEntries } from '@ansonlai/docx-redline-js';
import { inspectWordSource } from '../src/taskpane/modules/docx-redline-js-integration/consumer-core.js';
import * as wordOperationRunner from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';
import { computeWordRestorationToken, computeWordSourceRevisionToken, detectExportTailIndex } from '../src/taskpane/modules/docx-redline-js-integration/word-source-identity.js';
import { createDocumentInspectionSession } from '../src/taskpane/modules/chat/document-inspection-session.js';
import { createDocumentTools } from '../src/taskpane/modules/chat/document-tools.js';
import { createWordDocumentSourceReader, queueHostParagraphCount } from '../src/taskpane/modules/chat/word-document-source.js';
import { restoreBodyCheckpoint } from '../src/taskpane/modules/chat/checkpoint-restoration.js';
import { formatAutoCheckpointLabel, getCheckpoint, saveCheckpoint } from '../src/taskpane/modules/storage/checkpoint-store.js';

const status = message => { document.getElementById('status').textContent = message; };
function assert(condition, message) { if (!condition) throw new Error(message); }
function assertEqual(actual, expected, message) {
    assert(Object.is(actual, expected), `${message} Expected ${JSON.stringify(expected)}; received ${JSON.stringify(actual)}.`);
}
function assertDeepEqual(actual, expected, message) {
    assert(JSON.stringify(actual) === JSON.stringify(expected),
        `${message} Expected ${JSON.stringify(expected)}; received ${JSON.stringify(actual)}.`);
}
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
async function captureBodySource() {
    const run = instrumentedRun(createTransportCounters());
    return run(async context => {
        const body = context.document.body;
        const source = body.getOoxml();
        const readCount = queueHostParagraphCount(body);
        await context.sync();
        return { ooxml: source.value, hostParagraphCount: readCount() };
    });
}
function inspectBodySource(raw) {
    return { ...raw, ...inspectWordSource(raw.ooxml) };
}
function commentDurableIds(source) {
    const xml = source.parts?.commentsIdsXml
        || source.parts?.additionalParts?.['word/commentsIds.xml']
        || source.parts?.additionalParts?.['/word/commentsIds.xml']
        || '';
    const durableIds = new Map();
    for (const match of xml.matchAll(/<(?:[\w.-]+:)?commentId\b([^>]*)\/?\s*>/g)) {
        const paraId = match[1].match(/(?:[\w.-]+:)?paraId\s*=\s*"([^"]+)"/)?.[1];
        const durableId = match[1].match(/(?:[\w.-]+:)?durableId\s*=\s*"([^"]+)"/)?.[1];
        if (paraId && durableId) durableIds.set(paraId, durableId);
    }
    return durableIds;
}
function commentSnapshot(source) {
    const durableIds = commentDurableIds(source);
    return (source.inspection?.comments || []).map(comment => ({
        id: String(comment.id),
        paraId: comment.paraId || null,
        durableId: comment.paraId ? durableIds.get(comment.paraId) || null : null,
        author: String(comment.author || ''),
        date: comment.date || null,
        text: String(comment.text || ''),
        done: comment.done === true,
        parentCommentId: comment.parentCommentId == null ? null : String(comment.parentCommentId),
        parentParaId: comment.parentParaId || null,
        anchoredText: comment.anchoredText || '',
        targetRef: comment.targetRef || null
    }));
}
function assertCommentParentsConsistent(comments, label) {
    const byId = new Map(comments.map(comment => [comment.id, comment]));
    assert(byId.size === comments.length, `${label}: duplicate comment IDs were returned.`);
    for (const comment of comments) {
        if (comment.parentCommentId == null) continue;
        const parent = byId.get(comment.parentCommentId);
        assert(parent, `${label}: reply ${comment.id} lost parent ${comment.parentCommentId}.`);
        assert(comment.parentParaId === parent.paraId,
            `${label}: reply ${comment.id} points to a different parent paragraph.`);
    }
}
function assertPriorCommentsPreserved(before, after, label) {
    assert(after.length >= before.length, `${label}: comments were dropped.`);
    assertCommentParentsConsistent(after, label);
    const afterById = new Map(after.map(comment => [comment.id, comment]));
    const originalDurableIds = before.map(comment => comment.durableId);
    assert(originalDurableIds.every(Boolean), `${label}: source comments are missing durable IDs.`);
    assertEqual(new Set(originalDurableIds).size, originalDurableIds.length,
        `${label}: source durable IDs are not unique.`);
    for (const original of before) {
        const current = afterById.get(original.id);
        assert(current, `${label}: comment ${original.id} was removed or renumbered.`);
        for (const field of ['durableId', 'author', 'date', 'text', 'done', 'parentCommentId', 'anchoredText', 'targetRef']) {
            assertEqual(current[field], original[field], `${label}: comment ${original.id} changed ${field}.`);
        }
    }
}
function findAddedComment(comments, before, author, text, label) {
    const beforeDurableIds = new Set(before.map(comment => comment.durableId));
    const candidates = comments.filter(comment => comment.durableId && !beforeDurableIds.has(comment.durableId)
        && comment.author === author && comment.text === text);
    assertEqual(candidates.length, 1, `${label}: expected one new root comment by ${author}.`);
    const comment = candidates[0];
    assert(comment.durableId && comment.parentCommentId == null && comment.done === false,
        `${label}: new comment is missing a durable identity or was unexpectedly threaded/resolved.`);
    assertEqual(new Set(comments.map(item => item.durableId)).size, comments.length,
        `${label}: comment durable IDs are not unique.`);
    return comment;
}
function assertWorkflowCommentRetained(previous, comments, expectedAnchor, label) {
    assert(previous?.durableId, `${label}: the prior workflow comment has no durable identity.`);
    const matches = comments.filter(comment => comment.durableId === previous.durableId);
    assertEqual(matches.length, 1, `${label}: the workflow comment durable identity was dropped or duplicated.`);
    const current = matches[0];
    for (const field of ['author', 'date', 'text', 'done', 'targetRef']) {
        assertEqual(current[field], previous[field], `${label}: the workflow comment changed ${field}.`);
    }
    assertEqual(previous.parentCommentId, null, `${label}: expected the first workflow comment to be a root comment.`);
    assertEqual(current.parentCommentId, null, `${label}: the first workflow comment became a reply.`);
    assert(typeof expectedAnchor === 'string' && expectedAnchor.length > 0,
        `${label}: an explicit post-rewrite anchor expectation is required.`);
    assertEqual(current.anchoredText, expectedAnchor, `${label}: the workflow comment anchor changed unexpectedly.`);
    assertEqual(new Set(comments.map(comment => comment.durableId)).size, comments.length,
        `${label}: comment durable IDs are not unique.`);
    return {
        durableId: current.durableId,
        previousWordId: previous.id,
        currentWordId: current.id,
        idRemapped: previous.id !== current.id,
        anchoredText: current.anchoredText,
        targetRef: current.targetRef
    };
}
function revisionIdentity(source, author = null) {
    const xml = source.parts?.documentXml || '';
    const revisions = [...xml.matchAll(/<w:(ins|del)\b([^>]*?)(\/>|>([\s\S]*?)<\/w:\1>)/g)].map(match => {
        const attributes = match[2];
        const get = name => attributes.match(new RegExp(`(?:w:)?${name}\\s*=\\s*"([^"]*)"`))?.[1] || null;
        const body = match[4] || '';
        const text = [...body.matchAll(/<w:(?:t|delText)\b[^>]*>([\s\S]*?)<\/w:(?:t|delText)>/g)]
            .map(item => item[1].replaceAll('&amp;', '&').replaceAll('&lt;', '<').replaceAll('&gt;', '>')
                .replaceAll('&quot;', '"').replaceAll('&apos;', "'"))
            .join('');
        const paragraphIndex = [...xml.slice(0, match.index).matchAll(/<w:p\b/g)].length - 1;
        return { kind: match[1], id: get('id'), author: get('author'), date: get('date'), paragraphIndex, text, xml: match[0] };
    }).filter(revision => !author || revision.author === author);
    return revisions.sort((left, right) => JSON.stringify({ ...left, id: null, xml: null })
        .localeCompare(JSON.stringify({ ...right, id: null, xml: null })));
}
function sameRevisionRecords(left, right, includeIds) {
    const shape = revision => ({
        kind: revision.kind,
        ...(includeIds ? { id: revision.id } : {}),
        author: revision.author,
        date: revision.date,
        paragraphIndex: revision.paragraphIndex,
        text: revision.text
    });
    return [left.map(shape), right.map(shape)];
}
function revisionSemanticKey(revision) {
    return JSON.stringify({
        kind: revision.kind,
        author: revision.author,
        date: revision.date,
        paragraphIndex: revision.paragraphIndex,
        text: revision.text
    });
}
function assertRevisionRecordsPreserved(before, after, label, { allowIdRemap = false, author = null } = {}) {
    const expected = revisionIdentity(before, author);
    const actual = revisionIdentity(after, author);
    assert(expected.length > 0, `${label}: no source revisions were found.`);
    const [expectedShape, actualShape] = sameRevisionRecords(expected, actual, !allowIdRemap);
    assertDeepEqual(actualShape, expectedShape,
        `${label}: a pending revision's kind, author, date, text, or paragraph location changed.`);

    const actualByShape = new Map();
    for (const revision of actual) {
        const key = revisionSemanticKey(revision);
        const matches = actualByShape.get(key) || [];
        matches.push(revision);
        actualByShape.set(key, matches);
    }
    const idMappings = expected.map(revision => {
        const key = revisionSemanticKey(revision);
        const current = actualByShape.get(key)?.shift();
        assert(current, `${label}: could not map a source revision to the exported document.`);
        return {
            kind: revision.kind,
            author: revision.author,
            date: revision.date,
            paragraphIndex: revision.paragraphIndex,
            text: revision.text,
            checkpointId: revision.id,
            exportedId: current.id,
            idRemapped: revision.id !== current.id
        };
    });
    return { before: expected, after: actual, idMappings };
}
async function readWordUnderline(text) {
    return Word.run(async context => {
        const matches = context.document.body.search(text, { matchCase: true });
        matches.load('items/text');
        await context.sync();
        assert(matches.items.length === 1 && matches.items[0].text === text,
            `Expected one Word range for format readback text: ${text}`);
        matches.items[0].font.load('underline');
        await context.sync();
        return String(matches.items[0].font.underline);
    });
}
function firstDifference(left, right) {
    if (left === right) return null;
    let index = 0;
    while (index < left.length && index < right.length && left[index] === right[index]) index++;
    return index;
}
function sourcePartDiagnostics(expectedParts, restoredParts, hostParagraphCount) {
    const expectedEntries = extractDocumentPartsEntries(expectedParts);
    const restoredEntries = extractDocumentPartsEntries(restoredParts);
    const expectedByName = new Map(expectedEntries.map(entry => [entry.name, entry.payload]));
    const restoredByName = new Map(restoredEntries.map(entry => [entry.name, entry.payload]));
    const names = [...new Set([...expectedByName.keys(), ...restoredByName.keys()])].sort();
    const differences = names.filter(name => expectedByName.get(name) !== restoredByName.get(name)).map(name => {
        const expected = expectedByName.get(name);
        const restored = restoredByName.get(name);
        const index = typeof expected === 'string' && typeof restored === 'string' ? firstDifference(expected, restored) : null;
        return {
            name,
            expectedLength: typeof expected === 'string' ? expected.length : null,
            restoredLength: typeof restored === 'string' ? restored.length : null,
            firstDifferenceIndex: index,
            expectedSnippet: index === null || index < 0 ? null : expected.slice(Math.max(0, index - 70), index + 150),
            restoredSnippet: index === null || index < 0 ? null : restored.slice(Math.max(0, index - 70), index + 150),
        };
    });
    return { expectedParts: expectedEntries.map(entry => entry.name), restoredParts: restoredEntries.map(entry => entry.name), differences,
        expectedToken: null, restoredToken: null, hostParagraphCount };
}
async function post(route, body, type = 'application/json') {
    const response = await fetch(route, { method: 'POST', headers: { 'Content-Type': type }, body });
    if (!response.ok) throw new Error(`Local evidence server HTTP ${response.status}`);
}

async function createDirectDocumentTools(testCase, stats, options = {}) {
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
    let checkpointId = null;
    let checkpointOoxml = null;
    const receipts = [];
    const tools = createDocumentTools({
        session,
        getDocumentId: () => documentId,
        getAuthor: () => options.author || testCase.productionOptions?.author || 'WP4 reviewer',
        getGenerateRedlines: () => options.generateRedlines ?? testCase.productionOptions?.generateRedlines !== false,
        async persistCheckpoint(sourceOoxml) {
            checkpointCount++;
            checkpointOoxml = sourceOoxml;
            checkpointId = await saveCheckpoint(formatAutoCheckpointLabel('apply_document_edits'), sourceOoxml);
            await post(`/checkpoint/${testCase.name}`, sourceOoxml, 'application/xml');
            return checkpointId;
        },
        run,
        onReceipt: receipt => receipts.push(receipt),
    });
    return { tools, session, capturedSource, documentId, stats, receipts,
        get checkpointCount() { return checkpointCount; }, get checkpointId() { return checkpointId; },
        get checkpointOoxml() { return checkpointOoxml; } };
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
async function postWorkflowArtifact(testCase, stage) {
    await post(`/workflow-artifact/${testCase.name}/${stage}`, await exportDocx(), 'application/octet-stream');
}
async function runSequentialCommentWorkflow(testCase, harness) {
    const workflow = testCase.commentWorkflow;
    assert(workflow?.first && Array.isArray(workflow.secondBatch) && workflow.secondBatch.length === 2,
        `${testCase.name}: sequential comment workflow is malformed.`);
    const originalComments = commentSnapshot(harness.capturedSource);
    assertEqual(originalComments.length, workflow.expectedOriginalComments || 3,
        `${testCase.name}: source thread fixture does not contain the expected comments.`);
    assert(originalComments.some(comment => comment.parentCommentId != null)
        && originalComments.some(comment => comment.done),
    `${testCase.name}: source thread fixture lost its reply or resolved comment.`);

    const firstTarget = await handleFor(harness.tools, workflow.first.targetText);
    const firstReceipt = await harness.tools.applyDocumentEdits({ edits: [{
        type: 'comment', target: firstTarget, anchor: workflow.first.anchor, text: workflow.first.text
    }] });
    await postWorkflowArtifact(testCase, 'after-first-comment');
    assert(firstReceipt.success === true && firstReceipt.verified === true && firstReceipt.written === true
        && firstReceipt.writeAttempted === true && firstReceipt.accountedCount === 1,
    `${testCase.name}: first comment was not applied as a verified direct-tool write: ${JSON.stringify(firstReceipt)}`);
    assert(harness.checkpointCount === 1 && harness.stats.contentWrites === 1,
        `${testCase.name}: first comment did not persist one source checkpoint and write.`);

    const afterFirstSource = inspectBodySource(await captureBodySource());
    const afterFirstComments = commentSnapshot(afterFirstSource);
    assertEqual(afterFirstComments.length, originalComments.length + 1,
        `${testCase.name}: first comment did not add exactly one thread.`);
    assertPriorCommentsPreserved(originalComments, afterFirstComments, `${testCase.name} after first comment`);
    const firstAdded = findAddedComment(afterFirstComments, originalComments,
        testCase.productionOptions?.author || 'WP4 reviewer', workflow.first.text,
        `${testCase.name} first comment`);

    const handles = new Map();
    const edits = [];
    for (const proposal of workflow.secondBatch) {
        let handle = handles.get(proposal.targetText);
        if (!handle) {
            // The first write invalidated all old handles. This resolves the
            // second operation against a fresh source snapshot.
            handle = await handleFor(harness.tools, proposal.targetText);
            handles.set(proposal.targetText, handle);
        }
        const { targetText, ...edit } = proposal;
        edits.push({ ...edit, target: handle });
    }
    const secondReceipt = await harness.tools.applyDocumentEdits({ edits });
    await postWorkflowArtifact(testCase, 'after-second-comment-and-edit');
    const afterSecondSource = inspectBodySource(await captureBodySource());
    const afterSecondComments = commentSnapshot(afterSecondSource);
    assertPriorCommentsPreserved(originalComments, afterSecondComments, `${testCase.name} after second batch`);
    assertEqual(afterSecondComments.length, originalComments.length + 2,
        `${testCase.name}: second batch did not preserve the first comment and add exactly one more.`);
    const retainedFirstComment = assertWorkflowCommentRetained(firstAdded, afterSecondComments,
        workflow.expectedFirstAnchorAfterRewrite,
        `${testCase.name} first comment after second batch`);
    const secondAdded = findAddedComment(afterSecondComments, afterFirstComments,
        testCase.productionOptions?.author || 'WP4 reviewer', workflow.secondCommentText,
        `${testCase.name} second comment`);
    if (workflow.expectedFinalParagraph) {
        assert(afterSecondSource.inspection.paragraphs.some(paragraph => paragraph.exactText === workflow.expectedFinalParagraph),
            `${testCase.name}: combined text/comment batch did not produce the expected paragraph.`);
    }
    return {
        receipt: secondReceipt,
        originalComments,
        afterFirstComments,
        afterSecondComments,
        addedComments: [firstAdded, secondAdded],
        retainedFirstComment,
        firstReceipt,
        secondReceipt
    };
}
function assertCommitted(receipt, testCase, stats, checkpointCount, expectedContentWrites = 1, expectedCheckpoints = 1) {
    assert(receipt?.success === true && receipt.verified === true && receipt.completion === true,
        `${testCase.name}: document tools did not return a confirmed receipt: ${JSON.stringify(receipt)}`);
    assert(receipt.accountedCount === requestedEdits(testCase).length,
        `${testCase.name}: the receipt did not account for every requested edit: ${JSON.stringify(receipt)}`);
    assert(receipt.written === true && receipt.writeAttempted === true,
        `${testCase.name}: expected a single confirmed body write: ${JSON.stringify(receipt)}`);
    assert(Number.isInteger(receipt.checkpointIndex) && receipt.checkpointIndex > 0 && checkpointCount === expectedCheckpoints,
        `${testCase.name}: the same-source recovery checkpoint was not persisted.`);
    assert(receipt.counters?.contentWrites === 1 && receipt.counters?.verificationReads === 1
        && receipt.counters?.sourceReads >= 2,
    `${testCase.name}: unexpected direct-tool receipt counters: ${JSON.stringify(receipt.counters)}`);
    assert(stats.contentWrites === expectedContentWrites,
        `${testCase.name}: expected ${expectedContentWrites} OOXML body writes, got ${stats.contentWrites}.`);
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
    const applied = [];
    for (const step of preparation.filter(item => item.type === 'tracked-rewrite' || item.type === 'tracked-append')) {
        const previousMode = await getTrackingMode();
        await setTrackingMode(step.trackingMode || 'trackAll');
        try {
            const stats = createTransportCounters();
            const harness = await createDirectDocumentTools(testCase, stats, {
                generateRedlines: true,
                ...(step.author ? { author: step.author } : {})
            });
            let target, edit;
            if (step.type === 'tracked-append') {
                const context = await harness.tools.inspectDocument({ search: step.afterText || '', limit: 8, includeDocumentEnd: true });
                target = context.documentEnd?.handle;
                assert(context.status === 'ok' && target, `${testCase.name}: document-end handle is unavailable for tracked source preparation.`);
                edit = { type: 'append_paragraph', target, text: step.text };
            } else {
                target = await handleFor(harness.tools, step.targetText);
                edit = { type: 'rewrite_paragraph', target, text: step.text };
            }
            const receipt = await harness.tools.applyDocumentEdits({ edits: [edit] });
            assert(receipt.success === true && receipt.verified === true && receipt.written === true && receipt.writeAttempted === true,
                `${testCase.name}: tracked source preparation was not applied by the production document tool: ${JSON.stringify(receipt)}`);
            assert(harness.checkpointCount === 1 && stats.contentWrites === 1,
                `${testCase.name}: tracked source preparation did not persist one checkpoint and one source write.`);
            const prepared = await captureBodySource();
            const inspection = inspectWordSource(prepared.ooxml).inspection;
            assert(inspection.paragraphs.some(paragraph => paragraph.hasRevisions),
                `${testCase.name}: tracked source preparation did not leave a body revision in Word.`);
            applied.push({ type: step.type, targetText: step.targetText, text: step.text, author: step.author || 'WP4 reviewer', receipt,
                checkpointCount: harness.checkpointCount, contentWrites: stats.contentWrites,
                revisionedParagraphs: inspection.paragraphs.filter(paragraph => paragraph.hasRevisions).map(paragraph => paragraph.exactText) });
        } finally { await setTrackingMode(previousMode); }
    }
    const nativePreparation = preparation.filter(item => item.type !== 'tracked-rewrite' && item.type !== 'tracked-append');
    if (!nativePreparation.length) return applied;
    return [...applied, ...await Word.run(async context => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load('items/text');
        await context.sync();
        const nativeApplied = [];
        for (const step of nativePreparation) {
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
            nativeApplied.push({ type: step.type, paragraphText: step.paragraphText, level: step.level, numbering: step.numbering });
        }
        return nativeApplied;
    })];
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

async function runCheckpointRestoreControl(testCase, harness) {
    assert(Number.isInteger(harness.checkpointId) && harness.checkpointId > 0 && harness.checkpointOoxml,
        `${testCase.name}: the IndexedDB checkpoint identity is unavailable.`);
    const checkpoint = await getCheckpoint(harness.checkpointId);
    assert(checkpoint?.id === harness.checkpointId && checkpoint.ooxml === harness.checkpointOoxml,
        `${testCase.name}: IndexedDB did not return the exact persisted pre-edit checkpoint.`);
    const checkpointInspection = inspectWordSource(checkpoint.ooxml);
    const checkpointCommentCount = checkpointInspection.inspection.comments?.length || 0;
    const checkpointComments = commentSnapshot(checkpointInspection);
    const checkpointRevisionElementCount = (checkpointInspection.parts.documentXml.match(/<w:(?:ins|del)\b/g) || []).length;
    const changedText = testCase.agenticRequest?.text
        || testCase.restorationHandleText
        || requestedEdits(testCase).find(edit => typeof edit.text === 'string')?.text;
    assert(changedText, `${testCase.name}: restoration control needs a changed paragraph text.`);
    const changedHandle = await handleFor(harness.tools, changedText);
    const checkpointCountBeforeRestore = harness.checkpointCount;
    const contentWritesBeforeRestore = harness.stats.contentWrites;
    const trackingModeBefore = await getTrackingMode();
    const restoreStats = createTransportCounters();
    const receipt = await restoreBodyCheckpoint(checkpoint.ooxml, {
        run: instrumentedRun(restoreStats),
        getDocumentId: () => harness.documentId,
    });
    const trackingModeAfter = await getTrackingMode();
    const restoredSource = inspectBodySource(await captureBodySource());
    if (!receipt.success) {
        const expectedSource = inspectWordSource(checkpoint.ooxml);
        const diagnostics = sourcePartDiagnostics(expectedSource.parts, restoredSource.parts, restoredSource.hostParagraphCount);
        const expectedToken = await computeWordRestorationToken(expectedSource.parts, { hostParagraphCount: restoredSource.hostParagraphCount });
        const restoredToken = await computeWordRestorationToken(restoredSource.parts, { hostParagraphCount: restoredSource.hostParagraphCount });
        diagnostics.expectedToken = expectedToken;
        diagnostics.restoredToken = restoredToken;
        await post(`/restore-diagnostics/${testCase.name}`, JSON.stringify(diagnostics));
        await post(`/restored-source/${testCase.name}`, restoredSource.ooxml, 'application/xml');
        await post(`/restored/${testCase.name}`, await exportDocx(), 'application/octet-stream');
    }
    assert(receipt.success === true && receipt.verified === true && receipt.written === true && receipt.writeAttempted === true,
        `${testCase.name}: checkpoint restoration did not report a verified body write: ${JSON.stringify(receipt)}`);
    assert(trackingModeAfter === trackingModeBefore,
        `${testCase.name}: checkpoint restoration did not restore change tracking (${trackingModeBefore} → ${trackingModeAfter}).`);
    assert(restoreStats.contentWrites === 1 && restoreStats.nativeBodyTextWrites === 0
        && restoreStats.nativeParagraphInsertCalls.length === 0 && restoreStats.nativeListLevelWrites.length === 0
        && restoreStats.nativeListMethodCalls.length === 0,
    `${testCase.name}: checkpoint restoration used an unexpected host mutation route.`);
    const restoredComments = commentSnapshot(restoredSource);
    assertEqual(restoredComments.length, checkpointComments.length,
        `${testCase.name}: checkpoint restore did not recover its exact comment count.`);
    assertPriorCommentsPreserved(checkpointComments, restoredComments, `${testCase.name} checkpoint restore`);
    const staleHandleReceipt = await harness.tools.applyDocumentEdits({ edits: [{
        type: 'rewrite_paragraph', target: changedHandle, text: 'Must not survive checkpoint restoration.'
    }] });
    assert(staleHandleReceipt.success === false && staleHandleReceipt.written === false && staleHandleReceipt.writeAttempted === false
        && staleHandleReceipt.error?.code === 'STALE_DOCUMENT_CONTEXT',
    `${testCase.name}: a handle for the inserted paragraph remained usable after restoring its source checkpoint: ${JSON.stringify(staleHandleReceipt)}`);
    assert(harness.checkpointCount === checkpointCountBeforeRestore
        && harness.stats.contentWrites === contentWritesBeforeRestore,
        `${testCase.name}: stale post-restoration handle reached checkpoint persistence or a host content write.`);
    return {
        case: testCase.name,
        view: 'indexeddb-checkpoint-restoration',
        checkpointId: checkpoint.id,
        checkpointCommentCount,
        restoredCommentCount: restoredComments.length,
        checkpointRevisionElementCount,
        expectedRestoredBodyRevisions: Number.isInteger(testCase.expectedRestoredBodyRevisions)
            ? testCase.expectedRestoredBodyRevisions : null,
        exactPersistedXmlRead: true,
        receipt,
        trackingState: { before: trackingModeBefore, after: trackingModeAfter },
        postRestoreStaleHandle: staleHandleReceipt,
        hostTransport: {
            contentWrites: restoreStats.contentWrites,
            nativeBodyTextWrites: restoreStats.nativeBodyTextWrites,
            nativeParagraphInsertCalls: restoreStats.nativeParagraphInsertCalls,
            nativeListLevelWrites: restoreStats.nativeListLevelWrites,
            nativeListMethodCalls: restoreStats.nativeListMethodCalls,
        }
    };
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
        const restorationControls = [];
        const restorationCaseNames = new Set([
            'same-level-insert-redlines-disabled-restores-tracking',
            'direct-untracked-replacement-preserves-pending-revision',
            'direct-threaded-comments-sequential-merge-and-restore'
        ]);
        for (const testCase of cases) {
            status(`Running ${testCase.name}`);
            const fixture = await (await fetch(`/source/${testCase.name}`)).text();
            await seedFixture(fixture);
            const sourcePreparation = await prepareProductionSource(testCase);
            if (testCase.usesPreparedSource) {
                assert(sourcePreparation.length > 0, `${testCase.name}: prepared source was declared but no preparation ran`);
                await post(`/prepared-source/${testCase.name}`, await exportDocx(), 'application/octet-stream');
            }

            if (!testCase.expectedPrewriteRefusal && testCase.name === cases[0]?.name) {
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
            const formatBefore = testCase.formatReadback
                ? await readWordUnderline(testCase.formatReadback.text) : null;
            const commentWorkflowEvidence = testCase.commentWorkflow
                ? await runSequentialCommentWorkflow(testCase, harness) : null;
            const receipt = commentWorkflowEvidence?.receipt || await applyRequestedEdits(harness.tools, testCase);
            // Save the actual Word output before asserting its formatting or
            // revision identity, so a regression still leaves a reviewable package.
            await post(`/artifact/${testCase.name}`, await exportDocx(), 'application/octet-stream');
            const trackingModeAfter = await getTrackingMode();
            assert(trackingModeAfter === trackingModeBefore,
                `${testCase.name}: change-tracking state was not restored (${trackingModeBefore} → ${trackingModeAfter}).`);
            if (testCase.expectedPrewriteRefusal) {
                assert(receipt?.success === false && receipt?.error?.code === testCase.expectedPrewriteRefusal,
                    `${testCase.name}: expected prewrite refusal ${testCase.expectedPrewriteRefusal}, received ${JSON.stringify(receipt)}.`);
                assert(receipt.written === false && receipt.writeAttempted === false
                    && stats.contentWrites === 0 && harness.checkpointCount === 0,
                `${testCase.name}: unsupported operation reached checkpoint persistence or a host write.`);
                const afterSource = inspectBodySource(await captureBodySource());
                const beforeToken = await computeWordSourceRevisionToken(source.parts, {
                    hostParagraphCount: source.hostParagraphCount
                });
                const afterToken = await computeWordSourceRevisionToken(afterSource.parts, {
                    hostParagraphCount: afterSource.hostParagraphCount
                });
                assertDeepEqual(afterToken, beforeToken,
                    `${testCase.name}: prewrite refusal changed the canonical source.`);
                assertDeepEqual(revisionIdentity(afterSource), revisionIdentity(source),
                    `${testCase.name}: prewrite refusal changed pending revision identity.`);
                assertDeepEqual(commentSnapshot(afterSource), commentSnapshot(source),
                    `${testCase.name}: prewrite refusal changed comment threads.`);
                const refusalCheck = {
                    case: testCase.name,
                    view: 'direct-tool-prewrite-safety-refusal',
                    capability: 'unresolved-published-library-tracking-policy-limitation',
                    expectedCode: testCase.expectedPrewriteRefusal,
                    receipt,
                    sourceUnchanged: true,
                    hostTransport: {
                        contentWrites: stats.contentWrites,
                        checkpointCount: harness.checkpointCount
                    },
                    trackingState: { before: trackingModeBefore, after: trackingModeAfter }
                };
                checks.push(refusalCheck);
                continue;
            }
            assertCommitted(receipt, testCase, stats, harness.checkpointCount,
                commentWorkflowEvidence ? 2 : 1, commentWorkflowEvidence ? 2 : 1);
            let formatAfter = null;
            if (testCase.formatReadback) {
                formatAfter = await readWordUnderline(testCase.formatReadback.text);
                assertEqual(formatBefore, testCase.formatReadback.before,
                    `${testCase.name}: source formatting did not match the expected value.`);
                assertEqual(formatAfter, testCase.formatReadback.after,
                    `${testCase.name}: Word did not preserve the requested formatting after the direct write.`);
            }
            let revisionIdentityEvidence = null;
            if (testCase.revisionIdentityCheck) {
                const afterSource = inspectBodySource(await captureBodySource());
                const checkpointSource = inspectWordSource(harness.checkpointOoxml || '');
                const sourceRevisions = revisionIdentity(source);
                const checkpointRevisions = revisionIdentity(checkpointSource);
                const afterRevisions = revisionIdentity(afterSource);
                const [sourceIdentity, checkpointIdentity] = sameRevisionRecords(sourceRevisions, checkpointRevisions, true);
                const diagnostics = {
                    case: testCase.name,
                    foreignAuthor: testCase.revisionIdentityCheck.author,
                    source: sourceRevisions,
                    checkpoint: checkpointRevisions,
                    exported: afterRevisions
                };
                await post(`/revision-diagnostics/${testCase.name}`, JSON.stringify(diagnostics));
                assert(sourceRevisions.some(revision => revision.author === testCase.revisionIdentityCheck.author),
                    `${testCase.name}: no source revisions by ${testCase.revisionIdentityCheck.author} were found.`);
                assertDeepEqual(checkpointIdentity, sourceIdentity,
                    `${testCase.name}: checkpointing changed an existing revision's ID or identity.`);
                // Word may allocate a fresh revision ID while inserting a body
                // package. Existing revision meaning must still match exactly.
                revisionIdentityEvidence = assertRevisionRecordsPreserved(checkpointSource, afterSource,
                    `${testCase.name} exported source revisions`, { allowIdRemap: true });
            }
            const item = {
                case: testCase.name,
                view: 'actual-officejs-direct-document-tools',
                author: testCase.productionOptions?.author || 'WP4 reviewer',
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
            if (testCase.formatReadback) item.formatReadback = { before: formatBefore, after: formatAfter };
            if (testCase.revisionIdentityCheck) item.revisionIdentity = {
                author: testCase.revisionIdentityCheck.author,
                ...revisionIdentityEvidence
            };
            if (commentWorkflowEvidence) item.commentWorkflow = {
                firstReceipt: commentWorkflowEvidence.firstReceipt,
                secondReceipt: commentWorkflowEvidence.secondReceipt,
                originalComments: commentWorkflowEvidence.originalComments,
                afterFirstComments: commentWorkflowEvidence.afterFirstComments,
                afterSecondComments: commentWorkflowEvidence.afterSecondComments,
                addedComments: commentWorkflowEvidence.addedComments,
                retainedFirstComment: commentWorkflowEvidence.retainedFirstComment
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
            if (!testCase.expectedPrewriteRefusal && testCase.name === cases[0]?.name) {
                checks.push(...await runNoopAndRefusalControls(testCase, harness));
            }
            if (restorationCaseNames.has(testCase.name)) {
                const restorationControl = await runCheckpointRestoreControl(testCase, harness);
                await post(`/restored/${testCase.name}`, await exportDocx(), 'application/octet-stream');
                restorationControls.push(restorationControl);
                checks.push(restorationControl);
            }
        }
        await post('/result', JSON.stringify({
            status: 'passed', host: info, diagnostics: Office.context.diagnostics, checks, fixtureExpectations, restorationControls
        }));
        status('Direct document-tool transport checks passed. Exported documents require independent Word verification.');
    } catch (error) {
        status(`Failed: ${error.message}`);
        await post('/result', JSON.stringify({ status: 'failed', error: error.message })).catch(() => {});
    }
});
