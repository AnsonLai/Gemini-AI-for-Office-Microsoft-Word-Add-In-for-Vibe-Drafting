/* global Office, Word */
// Development-only golden scenario driver. Bundled ahead of the real taskpane
// only in `--env GOLDEN_SCENARIO=1` builds; ordinary builds never include it.
//
// It types each scenario prompt into the real chat UI and waits for the
// production chat loop to finish. In replay mode every Gemini request is
// answered by the local golden server; in live mode requests go to Gemini
// unchanged and paired with its response before the loop continues. After each step the document is exported
// for host-independent scoring.

const GEMINI_HOST = 'generativelanguage.googleapis.com';
const session = window.__GOLDEN_SESSION__;
const nativeFetch = globalThis.fetch.bind(globalThis);
const consoleLines = [];
let currentStep = null;
let pendingSourceDiagnostics = Promise.resolve();
let pendingModelLogs = Promise.resolve();

// Source snapshots can be enabled only by an explicit deterministic golden
// replay flag. The caller restricts this to the local NDA fixture and stores
// them under ignored .cache output; ordinary add-in builds have no driver.
if (session?.sourceDiagnostics === true) {
    globalThis.__WORD_SOURCE_DIAGNOSTICS__ = details => {
        if (!currentStep) return;
        const step = currentStep;
        pendingSourceDiagnostics = pendingSourceDiagnostics.then(() =>
            post('/golden/source-diagnostics', { step, details })
        );
    };
}

const stringify = value => {
    if (value instanceof Error) return `${value.name}: ${value.message}`;
    if (typeof value === 'string') return value;
    try { return JSON.stringify(value); } catch { return String(value); }
};
for (const level of ['log', 'info', 'warn', 'error']) {
    const original = console[level].bind(console);
    console[level] = (...args) => {
        consoleLines.push({ level, step: currentStep, message: args.map(stringify).join(' ').slice(0, 4000) });
        original(...args);
    };
}

async function post(route, body, type = 'application/json') {
    const response = await nativeFetch(route, {
        method: 'POST',
        headers: { 'Content-Type': type },
        body: type === 'application/json' ? JSON.stringify(body) : body
    });
    if (!response.ok) throw new Error(`Golden server ${route} HTTP ${response.status}`);
    return response;
}

// The request URL carries the API key: only the model name leaves this function.
const modelName = url => decodeURIComponent(String(url).split('/models/')[1]?.split(':')[0] || 'unknown');

function recordLiveExchange(exchange) {
    pendingModelLogs = pendingModelLogs.then(() => post('/golden/model-log', exchange));
    return pendingModelLogs;
}

globalThis.fetch = async (input, init = {}) => {
    const url = typeof input === 'string' ? input : input?.url;
    if (!session || !String(url).includes(GEMINI_HOST)) return nativeFetch(input, init);
    const payload = JSON.parse(init.body || '{}');
    const model = modelName(url);
    if (session.mode === 'replay') {
        const response = await post('/golden/model', { step: currentStep, model, payload });
        return new Response(await response.text(), { status: 200, headers: { 'Content-Type': 'application/json' } });
    }
    const step = currentStep;
    const started = Date.now();
    let response;
    try {
        response = await nativeFetch(input, init);
    } catch (error) {
        // Do not serialize fetch errors: some runtimes include the full URL,
        // which carries the API key. The boolean preserves the failed exchange.
        // Record only the kind of failure: an AbortError is the client's own timeout,
        // deadline or a cancel; anything else is a network failure (KB-13).
        await recordLiveExchange({ step, model, payload, status: null, response: null, requestError: true,
            requestErrorKind: error?.name === 'AbortError' ? 'aborted' : 'network', elapsedMs: Date.now() - started });
        throw error;
    }
    const responseText = await response.clone().text();
    await recordLiveExchange({ step, model, payload, status: response.status, response: responseText, elapsedMs: Date.now() - started });
    return response;
};

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
        const bytes = new Uint8Array(slices.reduce((total, slice) => total + slice.length, 0));
        let offset = 0;
        for (const slice of slices) { bytes.set(slice, offset); offset += slice.length; }
        return bytes;
    } finally { await officeCall(cb => file.closeAsync(cb)); }
}

const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

async function waitFor(predicate, timeoutMs, label) {
    const started = Date.now();
    while (!predicate()) {
        if (Date.now() - started > timeoutMs) throw new Error(`Timed out waiting for ${label}`);
        await delay(100);
    }
}

function chatMessages() {
    return Array.from(document.querySelectorAll('#chat-messages > *')).map(node => node.innerText.trim()).filter(Boolean);
}

async function runStep(step, index) {
    currentStep = step.id;
    const input = document.getElementById('chat-input');
    const send = document.getElementById(step.thinking ? 'think-button' : 'send-button');
    const before = chatMessages().length;
    const consoleStart = consoleLines.length;
    const startedAt = Date.now();

    input.value = step.prompt;
    send.click();
    // The production loop disables input while working and re-enables it when done.
    await waitFor(() => input.disabled, 5000, `step ${step.id} to start`);
    await waitFor(() => !input.disabled, session.stepTimeoutMs, `step ${step.id} to finish`);
    await delay(500);
    await pendingSourceDiagnostics;
    await pendingModelLogs;

    const transcript = chatMessages().slice(before);
    const bytes = await exportDocx();
    await post(`/golden/step/${index}/docx`, bytes, 'application/octet-stream');
    await post(`/golden/step/${index}/meta`, {
        id: step.id,
        durationMs: Date.now() - startedAt,
        transcript,
        console: consoleLines.slice(consoleStart)
    });
    currentStep = null;
}

Office.onReady(async info => {
    if (!session) return;
    if (info.host !== Office.HostType.Word) return;
    let phase = 'startup';
    try {
        // The taskpane wires its chat handlers in its own Office.onReady callback.
        phase = 'waiting for taskpane handlers';
        await waitFor(() => typeof document.getElementById('send-button')?.onclick === 'function', 30000, 'taskpane chat handlers');
        phase = 'claiming fixture';
        const claim = await post('/golden/claim', {});
        const { fixture } = await claim.json();
        // Seed this disposable document with the fixture (tracking off).
        phase = `loading fixture (${fixture?.length ?? 0} base64 chars)`;
        await Word.run(async context => {
            context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
            context.document.body.insertFileFromBase64(fixture, Word.InsertLocation.replace);
            await context.sync();
        });
        await delay(1000);
        phase = 'running steps';
        for (const [index, step] of session.steps.entries()) {
            await runStep(step, index);
        }
        await post('/golden/done', { status: 'completed', host: info, diagnostics: Office.context.diagnostics });
    } catch (error) {
        // Office errors carry the failing API and reason in code/debugInfo; message alone is often just "GeneralException".
        await post('/golden/done', { status: 'aborted', error: error.message, step: currentStep, phase,
            code: error.code, debugInfo: error.debugInfo, host: info, diagnostics: Office.context?.diagnostics,
            console: consoleLines.slice(-50) }).catch(() => {});
    }
});
