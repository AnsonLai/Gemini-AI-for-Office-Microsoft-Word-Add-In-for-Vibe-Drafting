/* global Office, Word */
// Development-only driver for the taskpane checks lane (scripts/run-taskpane-checks.mjs).
// Bundled ahead of the real taskpane only in `--env TASKPANE_CHECKS=1` builds.
//
// It automates the manual Word checks M1 to M14 of docs/plans/2026-11-07-state-machine-fixes.md
// inside the real taskpane in real Word: it drives the real DOM (clicks, values plus the
// events the handlers listen to) and uses Word.run only to prepare or read the document.
// Every Gemini request is answered here by a scripted responder; nothing leaves the machine.

import { restoreBodyCheckpoint } from '../../src/taskpane/modules/chat/checkpoint-restoration.js';
import { inspectWordSource } from '../../src/taskpane/modules/docx-redline-js-integration/consumer-core.js';
import { computeWordRestorationToken } from '../../src/taskpane/modules/docx-redline-js-integration/word-source-identity.js';

const GEMINI_HOST = 'generativelanguage.googleapis.com';
const session = window.__CHECKS_SESSION__;
const nativeFetch = globalThis.fetch.bind(globalThis);
const $ = id => document.getElementById(id);
const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
const DUMMY_KEY = 'dummy-checks-key-0000';

// ---------------------------------------------------------------- scripted model
const requests = [];          // every Gemini request: { kind, model, payload }
let chatQueue = [];           // steps for function-calling chat requests, consumed in order
let glanceQueue = [];         // optional steps for Glance requests; default answers every question
let unscripted = 0;

const resp = parts => ({ candidates: [{ content: { role: 'model', parts }, finishReason: 'STOP', index: 0 }] });
const say = (text, extra = {}) => ({ body: resp([{ text }]), ...extra });
const call = (name, args, extra = {}) => ({ body: resp([{ functionCall: { name, args } }]), ...extra });
const inspectStep = (search, extra) => call('inspect_document', { search, limit: 8 }, extra);
const emptyReply = { body: { candidates: [{ finishReason: 'SAFETY', safetyRatings: [] }] } };
const httpError = status => ({ status, body: { error: { message: 'scripted failure' } } });

function classify(payload) {
    const tools = Array.isArray(payload?.tools) ? payload.tools : [];
    return tools.some(tool => Array.isArray(tool?.functionDeclarations) || Array.isArray(tool?.function_declarations)) ? 'chat' : 'other';
}

// The newest inspect_document result in the request history gives the current edit handles.
function handleFor(payload, includes) {
    for (const message of [...(payload?.contents || [])].reverse()) {
        for (const part of message.parts || []) {
            if (part.functionResponse?.name !== 'inspect_document') continue;
            const text = (part.functionResponse.response?.content || []).map(item => item.text || '').join('');
            let result; try { result = JSON.parse(text); } catch { continue; }
            const target = (result.targets || []).find(item => item.handle && String(item.text).includes(includes));
            if (target) return target.handle;
        }
    }
    throw new Error(`no inspected handle for "${includes}" in the request history`);
}
const editStep = (build, extra) => payload => call('apply_document_edits', { edits: build(payload) }, extra);
const replaceStep = (find, replace, extra) => editStep(payload => [{ type: 'replace_text',
    target: handleFor(payload, find), find, replace }], extra);
const underlineStep = (find, extra) => editStep(payload => [{ type: 'format_text',
    target: handleFor(payload, find), find, properties: { underline: true } }], extra);
const PHRASE = session?.phrase || 'one (1) copy'; // text in the fixture that the scripted edits target
const EDIT_TEXT = `${PHRASE} ZQEDIT`; // the marker ZQEDIT in the body text means the edit is in the document
// inspect, one replace edit, closing text: the smallest request that edits the document
const editScript = (closing = 'Done.') => [inspectStep(PHRASE), replaceStep(PHRASE, EDIT_TEXT), say(closing)];

function sleep(ms, signal) {
    return new Promise((resolve, reject) => {
        const abort = () => { clearTimeout(timer); reject(new DOMException('Aborted', 'AbortError')); };
        const timer = setTimeout(() => { signal?.removeEventListener('abort', abort); resolve(); }, ms);
        if (signal?.aborted) abort(); else signal?.addEventListener('abort', abort, { once: true });
    });
}

globalThis.fetch = async (input, init = {}) => {
    const url = typeof input === 'string' ? input : input?.url;
    if (!session || !String(url).includes(GEMINI_HOST)) return nativeFetch(input, init);
    const payload = JSON.parse(init.body || '{}');
    const kind = classify(payload);
    const record = { kind, payload, model: decodeURIComponent(String(url).split('/models/')[1]?.split(':')[0] || '?'), at: Date.now() };
    requests.push(record);
    let step;
    if (kind === 'chat') {
        step = chatQueue.shift();
        if (!step) { unscripted++; step = say('Done.'); }
    } else {
        step = glanceQueue.shift() || (() => {
            const prompt = payload.contents?.[0]?.parts?.[0]?.text || '';
            const ids = [...prompt.matchAll(/\(ID: "([^"]+)"\)/g)].map(match => match[1]);
            return say(JSON.stringify(Object.fromEntries(ids.map(id => [id, { status: 'green', summary: `Fine (${id}).` }]))));
        })();
    }
    if (typeof step === 'function') step = step(payload, record);
    if (step.delay) await sleep(step.delay, init.signal);
    return new Response(JSON.stringify(step.body), { status: step.status || 200, headers: { 'Content-Type': 'application/json' } });
};

// ---------------------------------------------------------------- page helpers
async function post(route, body) {
    const response = await nativeFetch(route, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    if (!response.ok) throw new Error(`checks server ${route} HTTP ${response.status}`);
    return response;
}

async function until(predicate, timeoutMs, label) {
    const started = Date.now();
    while (!predicate()) {
        if (Date.now() - started > timeoutMs) throw new Error(`timed out waiting for ${label}`);
        await wait(25);
    }
}

const msgs = () => Array.from($('chat-messages').children);
const norm = text => String(text).replace(/\s+/g, ' ').trim();
const chatText = () => norm($('chat-messages').innerText);
function role(el) {
    return ['user', 'error', 'gemini', 'system'].find(name => el.classList.contains(name)) || 'other';
}
function buttonsOf(el) { return Array.from(el.querySelectorAll('button')).map(b => norm(b.innerText)).filter(Boolean); }
function dump(list = msgs()) {
    return list.map(el => `[${role(el)}] ${norm(el.innerText)}`).join(' | ');
}
// Long system messages are cut at 120 characters behind a "Show more" button.
function fullText(el) {
    const toggle = el.querySelector('.system-msg-toggle');
    if (toggle && toggle.innerText === 'Show more') toggle.click();
    return norm(el.innerText).replace(/ ?Show (less|more)/, '').replace(/ ?[↺↻] (Revert changes|Retry)$/, '');
}
const findButton = label => Array.from(document.querySelectorAll('#chat-messages button')).find(b => norm(b.innerText).includes(label));
const fire = (el, type) => el.dispatchEvent(new Event(type, { bubbles: true }));
const glanceSummaries = () => Array.from(document.querySelectorAll('.glance-summary')).map(el => norm(el.textContent));
const requestCount = kind => requests.filter(item => item.kind === kind).length;
const inputBusy = () => $('chat-input').disabled;

function sendMessage(text, type = 'fast') {
    $('chat-input').value = text;
    $(type === 'fast' ? 'send-button' : 'think-button').click();
}
async function sendAndWait(text, timeoutMs = 60000) {
    sendMessage(text);
    await until(inputBusy, 5000, 'the request to start');
    await until(() => !inputBusy(), timeoutMs, 'the request to finish');
    await wait(150);
}

// Arms a one-shot reaction to text appearing in the chat; the action runs synchronously
// inside the DOM mutation, so it lands during the step that produced the text.
function onceText(substr, action, timeoutMs = 30000) {
    return new Promise((resolve, reject) => {
        const root = $('chat-messages');
        let obs;
        const timer = setTimeout(() => { obs.disconnect(); reject(new Error(`"${substr}" never appeared in the chat`)); }, timeoutMs);
        obs = new MutationObserver(() => {
            if (!root.textContent.includes(substr)) return;
            obs.disconnect(); clearTimeout(timer);
            try { resolve(action()); } catch (error) { reject(error); }
        });
        obs.observe(root, { childList: true, subtree: true, characterData: true });
    });
}

async function openSettings() {
    $('settings-button').click();
    await until(() => $('settings-view').style.display === 'block', 3000, 'Settings to open');
    await wait(300);
}
async function closeSettings() {
    $('back-to-main').click();
    await until(() => $('main-view').style.display === 'block', 3000, 'the chat view to return');
    await wait(300);
}
async function glanceIdle(timeoutMs = 20000) {
    await until(() => !glanceSummaries().some(text => text === 'Checking...'), timeoutMs, 'Glance to stop checking');
}

// ---------------------------------------------------------------- Word helpers
let fixtureBase64 = null;
async function loadFixture() {
    await Word.run(async context => {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        context.document.body.insertFileFromBase64(fixtureBase64, Word.InsertLocation.replace);
        await context.sync();
    });
    await wait(600);
}
async function bodyText() {
    return Word.run(async context => { context.document.body.load('text'); await context.sync(); return context.document.body.text; });
}
async function appendTyped(text) {
    await Word.run(async context => {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        context.document.body.insertParagraph(text, Word.InsertLocation.end);
        await context.sync();
    });
}
async function titleUnderline() {
    return Word.run(async context => {
        const range = context.document.body.search('NON-DISCLOSURE AGREEMENT', { matchCase: true }).getFirst();
        range.font.load('underline'); await context.sync(); return range.font.underline;
    });
}

// ---------------------------------------------------------------- state reset
const glanceQuestions = (count = 2) => JSON.stringify(Array.from({ length: count }, (_, i) => ({
    id: `q${i + 1}`, title: `Check ${i + 1}`, question: `Is item ${i + 1} fine?` })));
function setStore(extra = {}) {
    const store = { geminiApiKey: DUMMY_KEY, geminiModelFast: 'gemini-flash-latest', geminiModelSlow: 'gemini-2.5-pro',
        redlineEnabled: 'true', redlineAuthor: 'Checks', glanceAutoRun: 'false', glanceSettings: glanceQuestions(2), ...extra };
    for (const key of ['geminiSystemMessage']) if (!(key in store)) localStorage.removeItem(key);
    for (const [key, value] of Object.entries(store)) {
        if (value === null) localStorage.removeItem(key); else localStorage.setItem(key, value);
    }
}

// Back to a known state: chat view, no request, empty chat, fresh document, default settings.
async function reset({ store = {}, document: reload = true } = {}) {
    if ($('settings-view').style.display === 'block') await closeSettings();
    if (inputBusy()) {
        document.querySelector('.cancel-request-btn')?.click();
        await until(() => !inputBusy(), 30000, 'the previous request to end');
    }
    await until(() => !$('refresh-chat-button').disabled, 30000, 'refresh to be available');
    setStore(store);
    // Settings round trip so the Glance cards are rebuilt from the stored questions.
    await openSettings(); await closeSettings();
    $('refresh-chat-button').click();
    await until(() => chatText().includes('Chat history cleared'), 5000, 'the chat to clear');
    if (reload) await loadFixture();
    chatQueue = []; glanceQueue = []; unscripted = 0; requests.length = 0;
    await glanceIdle();
}

// ---------------------------------------------------------------- test harness
let currentTester = null;
function tester() {
    const failed = [], passed = [], notes = [];
    return currentTester = {
        ok(condition, label, detail = '') { (condition ? passed : failed).push(detail ? `${label} (${detail})` : label); return condition; },
        note(text) { notes.push(text); },
        result(skipIfNothing = false) {
            const status = failed.length ? 'FAIL' : (skipIfNothing && !passed.length) ? 'SKIP' : 'PASS';
            const evidence = [...(failed.length ? [`FAILED: ${failed.join('; ')}`] : []),
                ...(passed.length ? [`ok: ${passed.join('; ')}`] : []), ...notes].join(' || ');
            return { status, evidence };
        },
    };
}
const skip = reason => ({ status: 'SKIP', evidence: reason });

// ---------------------------------------------------------------- the checks
const checks = [];
const check = (id, title, run) => checks.push({ id, title, run });

check('M1', 'Settings: dropping selected text on a Glance card does not reorder', async () => {
    const t = tester();
    await reset({ store: { glanceSettings: glanceQuestions(3) }, document: false });
    await openSettings();
    $('glance-settings-header').click();
    await wait(200);
    const order = () => Array.from(document.querySelectorAll('.glance-settings-card .glance-title-input')).map(input => input.value).join(',');
    const stored = () => JSON.parse(localStorage.getItem('glanceSettings')).map(item => item.title).join(',');
    const dragEvent = (type, target) => {
        const dataTransfer = new DataTransfer();
        target.dispatchEvent(new DragEvent(type, { bubbles: true, cancelable: true, dataTransfer }));
        return dataTransfer;
    };
    t.note('Events are synthetic DragEvents with a DataTransfer; real mouse dragging was not performed.');
    t.ok(order() === 'Check 1,Check 2,Check 3', 'start order', order());
    // Positive control: a real card drag by its handle (dragstart, drop, dragend) must still reorder.
    let cards = document.querySelectorAll('.glance-settings-card');
    dragEvent('dragstart', cards[0].querySelector('.drag-handle'));
    dragEvent('drop', cards[2]);
    dragEvent('dragend', cards[0].querySelector('.drag-handle'));
    t.ok(order() === 'Check 2,Check 3,Check 1', 'control: handle drag reorders', order());
    const before = `${order()} / stored ${stored()}`;
    // The bug case: text dragged out of a question box (no card dragstart) and dropped on cards.
    cards = document.querySelectorAll('.glance-settings-card');
    for (const index of [0, 1, 2]) {
        const dataTransfer = new DataTransfer();
        dataTransfer.setData('text/plain', 'selected words');
        cards[index].querySelector('.glance-question-input').dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer }));
    }
    t.ok(`${order()} / stored ${stored()}` === before, 'text drop on each card leaves the order unchanged', `${before} -> ${order()} / stored ${stored()}`);
    // A stale source must not survive: a finished drag followed by a text drop also changes nothing.
    cards = document.querySelectorAll('.glance-settings-card');
    dragEvent('dragstart', cards[1].querySelector('.drag-handle'));
    dragEvent('dragend', cards[1].querySelector('.drag-handle'));
    dragEvent('drop', cards[0]);
    t.ok(`${order()} / stored ${stored()}` === before, 'drop after a finished drag does not reorder', order());
    return t.result();
});

check('M2', 'Cancel right after sending shows grey "Request cancelled."', async () => {
    const t = tester();
    for (const variant of ['immediately after the send click', 'once the model request is in flight']) {
        await reset({ document: false });
        chatQueue = [inspectStep(PHRASE, { delay: 8000 })];
        sendMessage('M2 question about the document');
        await until(inputBusy, 5000, 'request to start');
        if (variant.startsWith('once')) await until(() => requestCount('chat') >= 1, 15000, 'the model request');
        const cancel = document.querySelector('.cancel-request-btn');
        t.ok(Boolean(cancel), `${variant}: Cancel button present`);
        cancel?.click();
        await until(() => !inputBusy(), 15000, 'the request to end');
        await wait(200);
        const last = msgs().at(-1);
        t.ok(last && role(last) === 'system' && norm(last.innerText) === 'System: Request cancelled.', `${variant}: grey system message`, dump(msgs().slice(-2)));
        t.ok(!msgs().some(el => role(el) === 'error'), `${variant}: no red error message`);
        t.ok(!document.querySelector('.cancel-request-btn'), `${variant}: Cancel button removed`);
    }
    return t.result();
});

check('M3', 'Obsolete: Save button removed', async () => skip('Obsolete: the Save button was removed in f1e74f4. Covered by M13.'));

check('M4', 'Glance without a key; results survive Settings; added question rebuilds cards', async () => {
    const t = tester();
    await reset({ document: true, store: { geminiApiKey: null } });
    $('refresh-glance-button').click();
    await wait(300);
    t.ok(glanceSummaries().every(text => text === 'Add your Gemini API key in Settings to run checks.'), 'no key: cards say to add a key', glanceSummaries().join(' / '));
    t.ok(requestCount('other') === 0, 'no key: no model request');
    setStore();
    $('refresh-glance-button').click();
    await glanceIdle();
    const results = glanceSummaries();
    t.ok(results.length === 2 && results.every(text => text.startsWith('Fine (')), 'with a key: Glance results shown', results.join(' / '));
    await openSettings(); await closeSettings();
    t.ok(JSON.stringify(glanceSummaries()) === JSON.stringify(results), 'results survive Settings then Back', glanceSummaries().join(' / '));
    t.ok(requestCount('other') === 1, 'Settings then Back ran nothing', `Glance requests ${requestCount('other')}`);
    await openSettings();
    $('add-glance-card-button').click();
    await closeSettings();
    const rebuilt = glanceSummaries();
    t.ok(rebuilt.length === 3 && rebuilt.every(text => text === 'Press refresh to run this check.'), 'added question: cards rebuilt idle', rebuilt.join(' / '));
    return t.result();
});

check('M5', 'One request that edits then checks ends with receipt and Revert', async () => {
    const t = tester();
    await reset();
    const startUnderline = await titleUnderline();
    chatQueue = [inspectStep('NON-DISCLOSURE'), underlineStep('NON-DISCLOSURE AGREEMENT'), inspectStep('NON-DISCLOSURE'),
        say('Confirmed: the title is underlined.')];
    await sendAndWait('M5 underline the title, then confirm it is underlined');
    const bubble = msgs().find(el => el.querySelector('.revert-checkpoint-btn'));
    t.ok(Boolean(bubble), 'a bubble with a Revert button remains', dump());
    t.ok(bubble && /Revert changes/.test(bubble.innerText), 'Revert button is the edit one');
    t.ok(bubble && !/Inspecting the document/.test(bubble.innerText), 'the later inspection did not replace the receipt', bubble && fullText(bubble));
    t.ok(msgs().some(el => role(el) === 'gemini' && /Confirmed/.test(el.innerText)), 'closing message shown');
    t.ok(requestCount('chat') === 4, 'four model turns', String(requestCount('chat')));
    const endUnderline = await titleUnderline();
    t.ok(endUnderline !== 'None' && startUnderline !== endUnderline, 'title underline really changed in Word', `${startUnderline} -> ${endUnderline}`);
    return t.result();
});

check('M6', 'Cancel while "Inspecting the document..." shows stops the request', async () => {
    const t = tester();
    await reset();
    chatQueue = [inspectStep(PHRASE), replaceStep(PHRASE, EDIT_TEXT), say('Done.')];
    const fired = onceText('Inspecting the document', () => { document.querySelector('.cancel-request-btn')?.click(); return true; });
    sendMessage('M6 inspect and edit the archival copies');
    t.ok(await fired, 'Cancel pressed while Inspecting text was showing');
    await until(() => !inputBusy(), 20000, 'the request to end');
    await wait(500);
    t.ok(msgs().some(el => /Request cancelled/.test(el.innerText)), 'cancel message shown', dump(msgs().slice(-2)));
    t.ok(requestCount('chat') === 1, 'no further model request after Cancel', `chat requests ${requestCount('chat')}`);
    t.ok(!(await bodyText()).includes('ZQEDIT'), 'no edit reached the document');
    t.ok(!inputBusy(), 'input unlocked');
    return t.result();
});

check('M7', 'Failure after tool run offers no Retry (unit tested)', async () => skip('Unit tested; cannot be triggered on demand.'));

check('M8', 'Refresh: refused during Revert; during an edit it cancels, waits, then clears', async () => {
    const t = tester();
    await reset();
    chatQueue = editScript();
    await sendAndWait('M8 change one copy to two copies');
    t.ok((await bodyText()).includes('ZQEDIT'), 'setup: edit landed');
    const revert = findButton('Revert changes');
    t.ok(Boolean(revert), 'setup: Revert button present');
    revert.click();
    $('refresh-chat-button').click();
    await wait(50);
    const refusal = chatText().includes('Wait for the revert to finish before clearing the chat.');
    t.ok(refusal, 'refresh during revert is refused with a message', chatText().slice(-160));
    t.ok(chatText().includes('Reverting to checkpoint') || chatText().includes('Reverted'), 'revert was running or done, chat not cleared yet');
    await until(() => chatText().includes('Reverted') || chatText().includes('could not'), 30000, 'revert to end');
    await wait(300);
    t.note(`Plain Revert (no later changes) ended with: ${fullText(msgs().at(-1))}`);
    t.ok(!(await bodyText()).includes('ZQEDIT'), 'revert restored the document');
    $('refresh-chat-button').click();
    await wait(300);
    t.ok(chatText().includes('Chat history cleared'), 'refresh works once the revert has ended');

    // Part B: refresh while "Validating the document edit batch..." shows.
    await reset();
    chatQueue = editScript();
    let snapshot;
    const fired = onceText('Validating the document edit batch', () => {
        $('refresh-chat-button').click();
        snapshot = { refreshDisabled: $('refresh-chat-button').disabled, inputDisabled: inputBusy(), oldChat: chatText().includes('M8B') };
        return true;
    });
    sendMessage('M8B change one copy to two copies');
    await fired;
    t.ok(snapshot.refreshDisabled, 'refresh button greys out at once');
    let inputEnabledWhileOldChatShown = false;
    const started = Date.now();
    while (chatText().includes('M8B') && Date.now() - started < 40000) {
        if (!inputBusy() && chatText().includes('M8B')) inputEnabledWhileOldChatShown = true;
        await wait(10);
    }
    t.ok(!chatText().includes('M8B') && chatText().includes('Chat history cleared'), 'chat clears after the request settles');
    t.ok(!inputEnabledWhileOldChatShown, 'input stayed locked until the chat cleared');
    t.ok(!$('refresh-chat-button').disabled, 'refresh button enabled again');
    await wait(500);
    const landed = (await bodyText()).includes('ZQEDIT');
    t.note(`Refresh at Validating text: edit ${landed ? 'LANDED' : 'did not land'}; fresh chat: ${dump()}`);
    t.ok(landed === Boolean(findButton('Revert changes')), 'landed-edit note and Revert present exactly when the edit landed');
    t.ok(requestCount('chat') <= 2, 'request stopped (no further model turns)', `chat requests ${requestCount('chat')}`);
    return t.result();
});

check('M9', '"Revert anyway" refused while a request runs and prompt kept', async () => {
    const t = tester();
    const variants = [
        ['text typed inside a paragraph', () => Word.run(async context => {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            context.document.body.search('NON-DISCLOSURE AGREEMENT', { matchCase: true }).getFirst().insertText(' M9TYPED', 'End');
            await context.sync();
        })],
        ['paragraph appended at the end', () => appendTyped('M9TYPED')],
    ];
    for (const [name, change] of variants) {
        await reset();
        chatQueue = editScript();
        await sendAndWait('M9 change one copy to two copies');
        await change();
        findButton('Revert changes').click();
        await until(() => findButton('Revert anyway'), 30000, 'the Revert anyway prompt');
        t.ok(true, `${name}: prompt shown after the document changed`, fullText(msgs().find(el => el.querySelector('.revert-btn-container') && /Revert anyway/.test(el.innerText))));
        chatQueue = [say('late answer', { delay: 9000 })];
        sendMessage('M9 second request');
        await until(inputBusy, 5000, 'second request to start');
        findButton('Revert anyway').click();
        await wait(300);
        t.ok(chatText().includes('Wait for the current request to finish, or cancel it, before reverting.'), `${name}: refusal message shown`, chatText().slice(-200));
        t.ok(Boolean(findButton('Revert anyway')) && Boolean(findButton('Keep document')), `${name}: prompt still there after the refusal`);
        document.querySelector('.cancel-request-btn')?.click();
        await until(() => !inputBusy(), 20000, 'second request to end');
        t.ok(Boolean(findButton('Revert anyway')), `${name}: prompt still there after the request ended`);
        findButton('Revert anyway').click();
        await until(() => !inputBusy() && !chatText().includes('Reverting to checkpoint...'), 30000, 'the revert');
        await wait(1000);
        const text = await bodyText();
        t.ok(!text.includes('ZQEDIT') && !text.includes('M9TYPED') && text.includes(PHRASE), `${name}: Revert anyway restores the pre-edit document`);
        const finalMessage = chatText().slice(-170);
        t.note(`${name}: Revert anyway ended with: ${finalMessage} (the verified-restore message is judged by R1)`);
    }
    return t.result();
});

check('M10', 'Failure after only reading vs failure after an applied edit', async () => {
    const t = tester();
    // A: failure after reading only
    await reset();
    chatQueue = [inspectStep(PHRASE), emptyReply, say('Done.')];
    sendMessage('M10A marker question');
    await until(inputBusy, 5000, 'start');
    await until(() => !inputBusy(), 30000, 'end');
    await wait(200);
    const errorEl = msgs().find(el => role(el) === 'error');
    t.ok(Boolean(errorEl), 'red error shown', dump());
    t.ok(!msgs().some(el => role(el) === 'system' && /Inspecting the document/.test(el.innerText)), 'bubble removed');
    t.ok(Boolean(errorEl) && buttonsOf(errorEl).some(label => label.includes('Retry')), 'Retry offered', errorEl && buttonsOf(errorEl).join(','));
    findButton('Retry').click();
    t.ok($('chat-input').value === 'M10A marker question', 'Retry puts the question back in the input (it does not send by itself)', $('chat-input').value);
    const before = requestCount('chat');
    sendMessage($('chat-input').value);
    await until(inputBusy, 5000, 'resend start');
    await until(() => !inputBusy(), 30000, 'resend end');
    const resent = requests.filter(item => item.kind === 'chat').slice(before);
    const userTurns = (resent[0]?.payload.contents || []).filter(turn => turn.role === 'user'
        && (turn.parts || []).some(part => String(part.text || '').includes('M10A marker question')));
    t.ok(userTurns.length === 1, 'resent request carries the question once, no duplicate', `user turns with the question: ${userTurns.length}`);
    t.ok((resent[0]?.payload.contents || []).length === 1, 'history held only the new prompt (failed attempt dropped)', `contents ${resent[0]?.payload.contents.length}`);
    t.ok(msgs().some(el => role(el) === 'gemini' && /Done\./.test(el.innerText)), 'retried request completed');

    // B1: empty reply after an applied edit; B2: HTTP 400 after an applied edit
    for (const [name, failure] of [['empty reply', emptyReply], ['HTTP 400', httpError(400)]]) {
        await reset();
        chatQueue = [inspectStep(PHRASE), replaceStep(PHRASE, EDIT_TEXT), failure];
        await sendAndWait(`M10B ${name} marker`);
        const bubble = msgs().find(el => el.querySelector('.revert-checkpoint-btn'));
        t.ok(Boolean(bubble), `${name} after edit: bubble keeps receipt and Revert`, dump());
        t.ok(!buttonsOf($('chat-messages')).some(label => label.includes('Retry')), `${name} after edit: no Retry`);
        const text = bubble ? fullText(bubble) : '';
        t.ok(text.endsWith('I no longer have the earlier conversation in memory, so restate anything I need.'), `${name} after edit: message ends with the memory note`, text.slice(-140));
        t.ok(!msgs().some(el => role(el) === 'error'), `${name} after edit: no separate red error`);
        t.ok((await bodyText()).includes('ZQEDIT'), `${name} after edit: the edit is in the document`);
    }
    return t.result();
});

check('M11', 'Refresh while a request cannot settle: 15 s wait, then message and lock kept', async () => {
    // Runs last: it deliberately leaves Word.run held until the end of the check.
    const t = tester();
    await reset();
    const realRun = Word.run.bind(Word);
    let release; let held = null; let heldCalls = 0;
    Word.run = (...args) => held ? (heldCalls++, held.then(() => realRun(...args))) : realRun(...args);
    try {
        chatQueue = editScript();
        const clearedBefore = (chatText().match(/Chat history cleared/g) || []).length;
        const fired = onceText('Validating the document edit batch', () => {
            held = new Promise(resolve => { release = resolve; });
            setTimeout(() => $('refresh-chat-button').click(), 300);
            return true;
        });
        sendMessage('M11 change one copy to two copies');
        await fired;
        t.note('Test double: Word.run calls made after the edit started are held by the driver (a host call that never returns); they are released at the end.');
        await wait(16500);
        if (!inputBusy() && heldCalls === 0) return skip('The edit did not make any held Word.run call, so the request still settled; could not simulate a request that never settles.');
        t.ok(heldCalls > 0, 'the request is stuck in a held Word.run call', `held calls ${heldCalls}`);
        t.ok(!$('refresh-chat-button').disabled, 'refresh button is back after 15 s');
        t.ok(chatText().includes('Word is still finishing the last request'), 'message "Word is still finishing the last request" shown', chatText().slice(-200));
        t.ok((chatText().match(/Chat history cleared/g) || []).length === clearedBefore, 'chat not cleared');
        t.ok(chatText().includes('M11 change'), 'old conversation still shown');
        t.ok(inputBusy(), 'input still locked');
        release(); held = null;
        await until(() => !inputBusy(), 40000, 'the released request to settle');
        $('refresh-chat-button').click();
        await until(() => !chatText().includes('M11 change'), 20000, 'refresh to clear after settling');
        t.ok(true, 'refresh clears the chat once the request settled', dump().slice(0, 200));
    } finally {
        Word.run = realRun; held = null; release?.();
    }
    return t.result();
});

check('M12', 'Refresh as the write lands: cleared note, landed-edit note with working Revert', async () => {
    const t = tester();
    // The closing model turn is slow, so a request is still running after the edit lands.
    const offsets = [150, 250, 350, 450, 600, 800, 1100];
    let landedSeen = 0; let notLandedSeen = 0; let finishedBeforeClick = 0;
    for (const offset of offsets) {
        if (landedSeen >= 1 && notLandedSeen >= 1) break;
        await reset();
        chatQueue = [inspectStep(PHRASE), replaceStep(PHRASE, EDIT_TEXT), say('Done.', { delay: 3000 })];
        let atClick;
        const fired = onceText('Validating the document edit batch', () => {
            setTimeout(() => { atClick = { busy: inputBusy() }; $('refresh-chat-button').click(); atClick.refreshDisabled = $('refresh-chat-button').disabled; }, offset);
            return true;
        });
        sendMessage('M12 change one copy to two copies');
        await fired;
        await wait(offset + 100);
        await until(() => !chatText().includes('M12 change') && !inputBusy(), 45000, 'refresh to finish');
        await wait(600);
        const landed = (await bodyText()).includes('ZQEDIT');
        const note = chatText().includes('An edit from the cancelled request was applied to the document.');
        const revert = findButton('Revert changes');
        const view = `refresh +${offset} ms after the Validating text: request ${atClick.busy ? 'still running' : 'already finished'} at the click, edit ${landed ? 'landed' : 'did not land'}, note ${note ? 'shown' : 'absent'}, Revert ${revert ? 'present' : 'absent'}`;
        t.note(view);
        if (!atClick.busy) { finishedBeforeClick++; continue; }
        t.ok(landed === note && note === Boolean(revert), `consistent: ${view}`, dump());
        if (landed) {
            landedSeen++;
            if (revert) {
                revert.click();
                await until(() => !inputBusy() && !chatText().includes('Reverting to checkpoint...'), 30000, 'revert to end');
                await wait(800);
                t.ok(!(await bodyText()).includes('ZQEDIT'), 'Revert in the landed-edit note restores the document', chatText().slice(-120));
            }
        } else {
            notLandedSeen++;
            const only = msgs().map(el => norm(el.innerText));
            t.ok(only.length === 2 && /Welcome/.test(only[0]) && /Chat history cleared/.test(only[1]), 'no landed edit: fresh chat shows only the first messages', only.join(' | '));
        }
    }
    // A refresh with no request at all.
    await reset({ document: false });
    $('refresh-chat-button').click();
    await wait(300);
    t.ok(!chatText().includes('cancelled request'), 'idle refresh shows no landed-edit note');
    const result = t.result();
    if (!landedSeen) {
        result.status = result.status === 'FAIL' ? 'FAIL' : 'PARTIAL';
        result.evidence += ' || Never hit the case where the refresh cancels a running request after its edit landed; only the no-landed-edit half was verified.';
    }
    return result;
});

check('M13', 'Settings: no Save button; fields save on change; key removal and re-entry', async () => {
    const t = tester();
    await reset({ document: false });
    await openSettings();
    const statusText = () => norm($('settings-status').textContent);
    t.ok(!Array.from(document.querySelectorAll('#settings-view button')).some(b => /^save$/i.test(norm(b.textContent))), 'no Save button in Settings');
    const fast = $('model-select-fast'), slow = $('model-select-slow');
    fast.value = 'gemini-2.5-flash'; fire(fast, 'change');
    t.ok(statusText() === 'Saved.' && localStorage.getItem('geminiModelFast') === 'gemini-2.5-flash', 'Fast model saves on change', statusText());
    slow.value = 'gemini-3.8-flash'; fire(slow, 'change');
    t.ok(statusText() === 'Saved.' && localStorage.getItem('geminiModelSlow') === 'gemini-3.8-flash', 'Slow model saves on change', statusText());
    const system = $('system-message-input');
    system.value = 'M13 custom system prompt'; fire(system, 'change');
    t.ok(statusText() === 'Saved.' && localStorage.getItem('geminiSystemMessage') === 'M13 custom system prompt', 'system prompt saves on change', statusText());
    const key = $('api-key-input');
    key.value = `${DUMMY_KEY}-two`; fire(key, 'change');
    t.ok(statusText() === 'Saved.' && localStorage.getItem('geminiApiKey') === `${DUMMY_KEY}-two`, 'API key saves on change', statusText());
    await closeSettings(); await openSettings();
    t.ok(fast.value === 'gemini-2.5-flash' && slow.value === 'gemini-3.8-flash' && system.value === 'M13 custom system prompt'
        && key.value === `${DUMMY_KEY}-two`, 'values survive Back and reopening', `${fast.value} / ${slow.value} / ${system.value} / key ${key.value.length} chars`);
    // Custom model id
    fast.value = 'custom'; fire(fast, 'change');
    const custom = $('model-custom-fast');
    custom.value = '../x'; fire(custom, 'change');
    t.ok(statusText().startsWith('Enter a model ID') && localStorage.getItem('geminiModelFast') === 'gemini-2.5-flash', 'custom id ../x shows an error and is not saved', `${statusText()} | stored ${localStorage.getItem('geminiModelFast')}`);
    custom.value = 'my-model-1'; fire(custom, 'change');
    t.ok(statusText() === 'Saved.' && localStorage.getItem('geminiModelFast') === 'my-model-1', 'a valid custom id is saved', statusText());
    // Clearing the key
    key.value = ''; fire(key, 'change');
    t.ok(statusText() === 'API key removed.' && localStorage.getItem('geminiApiKey') === null, 'clearing the key shows "API key removed." and removes it', statusText());
    await closeSettings();
    $('refresh-glance-button').click();
    await wait(300);
    t.ok(glanceSummaries().every(text => text === 'Add your Gemini API key in Settings to run checks.'), 'after Back, Glance refresh says to add a key', glanceSummaries().join(' / '));
    t.note('Back alone does not touch the idle Glance cards when auto-run is off; the message appears on refresh (as in M4).');
    await sendAndWait('M13 message without a key');
    t.ok(msgs().some(el => role(el) === 'error' && /Please set your Gemini API key/.test(el.innerText)), 'sending asks for a key', dump(msgs().slice(-2)));
    t.ok(requestCount('chat') === 0, 'no model request without a key');
    // Re-entering
    await openSettings();
    key.value = DUMMY_KEY; fire(key, 'change');
    t.ok(statusText() === 'Saved.', 'key re-entered and saved', statusText());
    await closeSettings();
    chatQueue = [say('M13 works again')];
    await sendAndWait('M13 message with a key again');
    t.ok(msgs().some(el => role(el) === 'gemini' && /M13 works again/.test(el.innerText)), 'the pane works again', dump(msgs().slice(-2)));
    $('refresh-glance-button').click();
    await glanceIdle();
    t.ok(glanceSummaries().every(text => text.startsWith('Fine (')), 'Glance works again', glanceSummaries().join(' / '));
    return t.result();
});

check('M14', 'Glance auto-run: changes then Back run once; Back without changes runs nothing', async () => {
    const t = tester();
    await reset({ document: true, store: { glanceAutoRun: 'true' } });
    t.ok(true, 'setup: auto-run enabled in storage');
    await openSettings();
    t.ok($('glance-autorun-toggle').checked, 'Settings shows auto-run on');
    const settle = async () => { await wait(1200); await glanceIdle(); await wait(1200); };
    const question = document.querySelector('.glance-question-input');
    question.value = 'M14 changed question?'; fire(question, 'change');
    await closeSettings(); await settle();
    t.ok(requestCount('other') === 1, 'changed question then Back: Glance ran once', `Glance requests ${requestCount('other')}`);
    const results = glanceSummaries();
    t.ok(results.every(text => text.startsWith('Fine (')), 'results filled in', results.join(' / '));
    await openSettings(); await closeSettings(); await settle();
    t.ok(requestCount('other') === 1, 'Back with no changes: nothing re-ran', `Glance requests ${requestCount('other')}`);
    t.ok(JSON.stringify(glanceSummaries()) === JSON.stringify(results), 'Back with no changes: results kept', glanceSummaries().join(' / '));
    await openSettings();
    const fast = $('model-select-fast');
    fast.value = 'gemini-2.5-flash'; fire(fast, 'change');
    await closeSettings(); await settle();
    t.ok(requestCount('other') === 2, 'changed model then Back: Glance ran once more', `Glance requests ${requestCount('other')}, model ${requests.at(-1)?.model}`);
    t.ok(requests.at(-1)?.model === 'gemini-2.5-flash', 'the run used the new Fast model', requests.at(-1)?.model);
    return t.result();
});

// Not in the plan: the plain Revert path (no later changes) as a control for M8, M9 and M12.
check('R1', 'Control: plain Revert after an edit reports a verified restore', async () => {
    const t = tester();
    for (const redline of ['true', 'false']) {
        await reset({ store: { redlineEnabled: redline } });
        chatQueue = editScript();
        await sendAndWait(`R1 change the text (redlines ${redline})`);
        const landed = (await bodyText()).includes('ZQEDIT');
        findButton('Revert changes').click();
        await until(() => !inputBusy() && !chatText().includes('Reverting to checkpoint...') && /Reverted|could not|not available/.test(chatText()), 30000, 'revert to end');
        await wait(800);
        const text = await bodyText();
        const message = norm(msgs().at(-1).innerText);
        t.ok(landed, `redlines ${redline}: edit landed`);
        t.ok(!text.includes('ZQEDIT') && text.includes(PHRASE), `redlines ${redline}: document text restored`);
        t.ok(/Reverted and verified/.test(message), `redlines ${redline}: Revert reports a verified restore`, message);
    }
    return t.result();
});

// All stored checkpoint records, oldest first (read straight from the add-in's IndexedDB).
async function latestCheckpoints() {
    const [info] = (await indexedDB.databases()).filter(item => /checkpoint/i.test(item.name || ''));
    const db = await new Promise((resolve, reject) => { const req = indexedDB.open(info.name); req.onsuccess = () => resolve(req.result); req.onerror = () => reject(req.error); });
    try {
        const store = db.objectStoreNames[0];
        return (await new Promise((resolve, reject) => { const req = db.transaction(store).objectStore(store).getAll(); req.onsuccess = () => resolve(req.result); req.onerror = () => reject(req.error); }))
            .sort((a, b) => a.id - b.id);
    } finally { db.close(); }
}
// Which canonical parts differ between a checkpoint and the body Word exports after restoring it,
// ignoring the ids Word regenerates (paragraph/text ids, rsids).
async function restoreDifferences(checkpointOoxml, restoredOoxml) {
    const current = restoredOoxml || await Word.run(async context => { const ooxml = context.document.body.getOoxml(); await context.sync(); return ooxml.value; });
    const flatten = parts => {
        const flat = {};
        for (const [name, value] of Object.entries(parts || {})) {
            if (name === 'additionalParts' && value && typeof value === 'object') {
                for (const [part, payload] of Object.entries(value)) flat[part] = typeof payload === 'string' ? payload : JSON.stringify(payload);
            } else flat[name] = typeof value === 'string' ? value : JSON.stringify(value);
        }
        // Strip only what the restoration token also ignores, so anything left is what fails verification.
        for (const name of Object.keys(flat)) {
            flat[name] = String(flat[name]).replace(/ w:(?:rsidR|rsidRDefault|rsidRPr|rsidP|rsidDel|rsidSect|rsidTr)="[^"]*"|<w:rsids>.*?<\/w:rsids>/g, '');
            if (/^documentXml$|footnotes|endnotes|comments/i.test(name)) {
                flat[name] = flat[name].replace(/<w:p(?=[ >])[^>]*>/g, tag => tag.replace(/ w14:(?:paraId|textId)="[^"]*"/g, ''));
            }
        }
        return flat;
    };
    const before = flatten(inspectWordSource(checkpointOoxml).parts), after = flatten(inspectWordSource(current).parts);
    return [...new Set([...Object.keys(before), ...Object.keys(after)])].filter(name => before[name] !== after[name]).map(name => {
        const a = before[name] ?? '', b = after[name] ?? '';
        let i = 0;
        while (i < a.length && a[i] === b[i]) i++;
        return { part: name, checkpointLength: a.length, restoredLength: b.length, at: i,
            checkpoint: a.slice(Math.max(0, i - 80), i + 160), restored: b.slice(Math.max(0, i - 80), i + 160) };
    });
}
// Diagnostic for R1: call the production restore function directly, with and without the Revert guard.
check('R2', 'Diagnostic: restoreBodyCheckpoint called directly, with and without the guard token', async () => {
    const t = tester();
    for (const withGuard of [false, true]) {
        await reset();
        chatQueue = editScript();
        await sendAndWait(`R2 change the text (guard ${withGuard})`);
        let record = null;
        const all = await latestCheckpoints();
        record = all.at(-1) || null;
        if (!record) { t.ok(false, 'a stored checkpoint was found'); continue; }
        // Capture the body in the same Word.run as the verification read, then again later.
        let atVerify = null;
        const run = callback => Word.run(async context => {
            const result = await callback(context);
            const ooxml = context.document.body.getOoxml(); context.document.body.paragraphs.load('items'); await context.sync();
            atVerify = { ooxml: ooxml.value, count: context.document.body.paragraphs.items.length };
            return result;
        });
        const outcome = await restoreBodyCheckpoint(record.ooxml, { run, getDocumentId: () => 'r2', ...(withGuard ? { guardToken: record.afterToken } : {}) });
        t.ok(outcome.success === true, `direct restore ${withGuard ? 'with' : 'without'} the guard token verifies`, `checkpoint ${record.id}: ${outcome.message}`);
        if (!outcome.success && !withGuard) {
            const token = async (ooxml, options) => { try { const value = await computeWordRestorationToken(inspectWordSource(ooxml).parts, options); return `${value.value.slice(0, 8)}/${JSON.stringify(value.commentIdentity).length}`; } catch (error) { return `threw ${error.message}`; } };
            t.note(`tokens: checkpoint ${await token(record.ooxml)}, at verify ${await token(atVerify.ooxml, { hostParagraphCount: atVerify.count })} (no count ${await token(atVerify.ooxml)}, host paragraphs ${atVerify.count})`);
            t.note(`differences at verify: ${JSON.stringify(await restoreDifferences(record.ooxml, atVerify.ooxml))}`);
            t.note(`differences later: ${JSON.stringify(await restoreDifferences(record.ooxml))}`);
        }
    }
    return t.result();
});

// M11 leaves the pane in an unusual state, so it runs last.
const order = ['M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M12', 'M13', 'M14', 'R1', 'R2', 'M11'];
checks.sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id));

Office.onReady(async info => {
    if (!session || info.host !== Office.HostType.Word) return;
    let phase = 'startup';
    try {
        phase = 'waiting for taskpane handlers';
        await until(() => typeof $('send-button')?.onclick === 'function', 30000, 'taskpane chat handlers');
        phase = 'claiming fixture';
        fixtureBase64 = (await (await post('/checks/claim', {})).json()).fixture;
        phase = 'loading fixture';
        await loadFixture();
        for (const item of checks) {
            if (session.only && !session.only.includes(item.id)) continue;
            phase = item.id;
            let result;
            const started = Date.now();
            currentTester = null;
            try { result = await item.run(); }
            catch (error) { result = { status: 'FAIL', evidence: `harness error: ${error.message} || chat: ${dump().slice(0, 400)} || so far: ${currentTester?.result().evidence}` }; }
            await post('/checks/result', { id: item.id, title: item.title, ...result, ms: Date.now() - started });
        }
        await post('/checks/done', { status: 'completed' });
    } catch (error) {
        await post('/checks/done', { status: 'aborted', error: error.message, phase }).catch(() => {});
    }
});
