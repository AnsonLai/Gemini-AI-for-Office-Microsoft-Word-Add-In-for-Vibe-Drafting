// Golden scenario runner: drives the real taskpane in desktop Word through the
// owner's 16-prompt NDA session and scores every step from exported documents.
//
//   node scripts/run-golden-scenario.mjs --launch                 # deterministic replay
//   $env:GEMINI_API_KEY='...'; node scripts/run-golden-scenario.mjs --mode live --launch
//   node scripts/run-golden-scenario.mjs --cleanup                # remove the sideload registration
//
// Uses existing localhost development certificates; never installs trust.
// Serves on port 3001 so the developer's dev server (3000) and its stored
// settings are untouched. See scripts/README.md "Golden scenario".
import '../tests/setup-xml-provider.mjs';
import { createServer } from 'node:https';
import { execFileSync, spawn } from 'node:child_process';
import { appendFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { homedir } from 'node:os';
import { dirname, extname, join, resolve } from 'node:path';
import { scenario } from './golden/nda-scenario.mjs';
import { authorAttribution, buildDocumentModel, scoreStep } from './golden/golden-checks.mjs';
import { createReplayQueue } from './golden/golden-replay.mjs';

const require = createRequire(import.meta.url);
const root = resolve(import.meta.dirname, '..');
const PORT = 3001;
const ADDIN_ID = 'a81aa9d8-d01b-4196-b79d-e4b92030fc90';

function argument(name, fallback = null) {
    const index = process.argv.indexOf(name);
    if (index < 0) return fallback;
    const value = process.argv[index + 1];
    if (!value || value.startsWith('--')) throw new Error(`Missing value for ${name}`);
    return value;
}
const flag = name => process.argv.includes(name);

const stamp = new Date().toISOString().replace(/[:.]/g, '-');
const artifacts = resolve(root, argument('--artifacts-dir', `.cache/golden/${stamp}`));
const distDirectory = resolve(root, argument('--dist-dir', '.cache/golden/dist'));
const manifestPath = resolve(root, '.cache/golden/golden-manifest.xml');

if (flag('--cleanup')) {
    if (existsSync(manifestPath)) await require('office-addin-dev-settings').unregisterAddIn(manifestPath);
    console.log('Golden scenario add-in registration removed.');
    process.exit(0);
}

const mode = argument('--mode', 'replay');
if (!['replay', 'live'].includes(mode)) throw new Error('--mode must be replay or live');
const apiKey = mode === 'live' ? process.env.GEMINI_API_KEY : 'golden-replay-key';
if (!apiKey) throw new Error('Live mode requires GEMINI_API_KEY in the environment.');
const fastModel = argument('--model', null);
const stepTimeoutMs = Number(argument('--step-timeout', mode === 'live' ? '300' : '120')) * 1000;
const lastStep = Number(argument('--until', String(scenario.steps.length)));
const steps = scenario.steps.slice(0, lastStep);

mkdirSync(artifacts, { recursive: true });
mkdirSync(dirname(manifestPath), { recursive: true });

if (!flag('--skip-build')) {
    console.log('Building golden taskpane bundle...');
    execFileSync(process.execPath, [require.resolve('webpack/bin/webpack.js'), '--mode', 'development',
        '--env', 'GOLDEN_SCENARIO=1', '--output-path', distDirectory], { cwd: root, stdio: 'inherit' });
}

writeFileSync(manifestPath, readFileSync(join(root, 'manifest.xml'), 'utf8')
    .replaceAll('3dcfdb34-70c3-4bfe-8d5e-85089afcf673', ADDIN_ID)
    .replaceAll('Gemini AI for Office', 'Golden scenario (local)')
    .replace(/(id="TaskpaneButton\.Label" DefaultValue=")Assistant"/g, '$1Golden"')
    .replaceAll('https://localhost:3000/', `https://localhost:${PORT}/`));

// Settings live in the 3001 origin's storage, isolated from the real add-in.
const settings = {
    geminiApiKey: apiKey,
    redlineEnabled: 'true',
    redlineAuthor: 'Golden Scenario',
    glanceSettings: '[]',
    ...(fastModel ? { geminiModelFast: fastModel } : {})
};
const sessionScript = `<script>(function(){var s=${JSON.stringify(settings)};for(var k in s){try{localStorage.setItem(k,s[k]);}catch(e){}}
window.__GOLDEN_SESSION__=${JSON.stringify({ mode, stepTimeoutMs, steps: steps.map(({ id, prompt, thinking }) => ({ id, prompt, thinking: Boolean(thinking) })) })};})();</script>`;

const report = { scenario: scenario.name, mode, model: fastModel, startedAt: new Date().toISOString(), status: 'running', steps: [] };
const passedChecks = new Set();
const queues = new Map(steps.map(step => [step.id, createReplayQueue(step)]));
const modelLog = join(artifacts, 'model-log.jsonl');
const pendingDocx = new Map();
let claimed = false;
let launchedWordPid = null;
let previousAuthors = {};

async function readBody(request, limit = 50 * 1024 * 1024) {
    const parts = [];
    let size = 0;
    for await (const part of request) {
        size += part.length;
        if (size > limit) throw new Error('Request body too large');
        parts.push(part);
    }
    return Buffer.concat(parts);
}

const icon = result => result.ok ? 'PASS' : result.knownIssue ? 'KNOWN' : 'FAIL';

async function scoreCompletedStep(index, meta) {
    const step = steps[index];
    const bytes = pendingDocx.get(index);
    const file = `step-${String(index + 1).padStart(2, '0')}-${step.id}.docx`;
    writeFileSync(join(artifacts, file), bytes);
    let results;
    try {
        const model = await buildDocumentModel(bytes);
        results = scoreStep(model, steps, index, passedChecks);
        const attribution = authorAttribution(previousAuthors, model, settings.redlineAuthor);
        results.push({ key: `${step.id}#author`, step: step.id, own: true, label: 'redlines use the configured author', ...attribution,
            ...(step.authorKnownIssue ? { knownIssue: step.authorKnownIssue } : {}) });
        previousAuthors = model.revisionAuthors;
    } catch (error) {
        results = [{ key: `${step.id}#model`, step: step.id, own: true, label: 'document model', ok: false, detail: error.message }];
    }
    for (const result of results) if (result.own && result.ok) passedChecks.add(result.key);
    const replay = mode === 'replay' ? queues.get(step.id).summary() : null;
    const replayProblems = replay ? [...replay.desyncs, ...(replay.remaining ? [`${replay.remaining} scripted response(s) unused`] : [])] : [];
    const entry = {
        id: step.id,
        prompt: step.prompt,
        knownIssue: step.knownIssue || null,
        durationMs: meta.durationMs,
        docx: file,
        own: results.filter(result => result.own),
        regressions: results.filter(result => !result.own && !result.ok),
        replayProblems,
        transcript: meta.transcript,
        consoleErrors: meta.console.filter(line => line.level === 'error' || line.level === 'warn').map(line => line.message)
    };
    report.steps.push(entry);
    writeFileSync(join(artifacts, `step-${String(index + 1).padStart(2, '0')}-${step.id}.json`), JSON.stringify({ ...entry, console: meta.console }, null, 2));

    console.log(`\n[${index + 1}/${steps.length}] ${step.id} (${Math.round(meta.durationMs / 1000)}s)`);
    for (const result of entry.own) console.log(`  ${icon(result)} ${result.label}${result.ok ? '' : ` — ${result.detail}`}`);
    for (const result of entry.regressions) console.log(`  REGRESSION ${result.step}: ${result.label} — ${result.detail}`);
    for (const problem of replayProblems) console.log(`  REPLAY ${problem}`);
    if (entry.transcript.length) console.log(`  chat: ${entry.transcript.at(-1).replace(/\s+/g, ' ').slice(0, 200)}`);
}

function finish(summary) {
    const ownFailures = report.steps.flatMap(step => step.own.filter(result => !result.ok && !result.knownIssue));
    const knownFailures = report.steps.flatMap(step => step.own.filter(result => !result.ok && result.knownIssue));
    const regressions = report.steps.flatMap(step => step.regressions);
    const replayProblems = report.steps.flatMap(step => step.replayProblems);
    const complete = summary.status === 'completed' && report.steps.length === steps.length;
    report.status = complete && !ownFailures.length && !regressions.length && !replayProblems.length ? 'passed' : 'failed';
    report.finishedAt = new Date().toISOString();
    report.host = summary.host || null;
    if (summary.status !== 'completed') report.abort = { error: summary.error, step: summary.step, console: summary.console };
    report.totals = {
        steps: `${report.steps.length}/${steps.length}`,
        checks: report.steps.reduce((total, step) => total + step.own.length, 0),
        failed: ownFailures.length,
        knownIssueFailures: knownFailures.length,
        regressions: regressions.length,
        replayProblems: replayProblems.length
    };
    writeFileSync(join(artifacts, 'report.json'), JSON.stringify(report, null, 2));
    writeFileSync(join(artifacts, 'report.md'), renderMarkdown(report));
    console.log(`\nGolden scenario ${report.status.toUpperCase()} (${mode}): ${JSON.stringify(report.totals)}`);
    if (report.abort) console.log(`Aborted at ${report.abort.step}: ${report.abort.error}`);
    console.log(`Artifacts: ${artifacts}`);
    process.exitCode = report.status === 'passed' ? 0 : 1;
    // Close only the disposable Word instance this run started (its document is a temp file).
    if (launchedWordPid && !flag('--keep-word')) {
        try { process.kill(launchedWordPid); } catch { /* already closed */ }
    }
    // Remove the developer registration so ordinary Word sessions never show the golden add-in.
    const unregister = launchedWordPid && !flag('--keep-word')
        ? require('office-addin-dev-settings').unregisterAddIn(manifestPath).catch(error => console.error(`Unregister failed: ${error.message}`))
        : Promise.resolve();
    unregister.finally(() => setTimeout(() => server.close(() => process.exit()), 500));
}

function renderMarkdown(data) {
    const lines = [`# Golden scenario: ${data.scenario} (${data.mode}${data.model ? `, ${data.model}` : ''})`, '',
        `Status: **${data.status}** — ${JSON.stringify(data.totals)}`, `Started ${data.startedAt}, finished ${data.finishedAt}.`, ''];
    if (data.abort) lines.push(`Aborted at \`${data.abort.step}\`: ${data.abort.error}`, '');
    lines.push('| Step | Result | Checks |', '| --- | --- | --- |');
    for (const step of data.steps) {
        const failed = step.own.filter(result => !result.ok);
        const unexpected = failed.filter(result => !result.knownIssue);
        const result = failed.length === 0 && !step.regressions.length && !step.replayProblems.length ? 'pass'
            : !unexpected.length && !step.regressions.length && !step.replayProblems.length ? 'known issue' : 'FAIL';
        const checks = step.own.map(check => `${check.ok ? '✓' : check.knownIssue ? '◐' : '✗'} ${check.label}${check.ok ? ''
            : `: ${check.detail}${check.knownIssue ? ` ([known issue](../../../${check.knownIssue}))` : ''}`}`)
            .concat(step.regressions.map(check => `✗ regression of ${check.step} ${check.label}: ${check.detail}`))
            .concat(step.replayProblems.map(problem => `✗ replay: ${problem}`));
        lines.push(`| ${step.id} | ${result} | ${checks.join('<br>').replace(/\|/g, '\\|')} |`);
    }
    return `${lines.join('\n')}\n`;
}

const certs = join(homedir(), '.office-addin-dev-certs');
const tls = { cert: readFileSync(join(certs, 'localhost.crt')), key: readFileSync(join(certs, 'localhost.key')) };
async function handle(request, response) {
    try {
        const path = new URL(request.url, `https://localhost:${PORT}`).pathname;
        response.setHeader('Cache-Control', 'no-store');
        const json = value => { response.setHeader('Content-Type', 'application/json'); response.end(JSON.stringify(value)); };

        if (path === '/golden/claim' && request.method === 'POST') {
            await readBody(request);
            if (claimed) { response.writeHead(409).end(); return; }
            claimed = true;
            console.log(`Run claimed by Word; ${steps.length} step(s), ${mode} mode.`);
            json({ fixture: readFileSync(resolve(root, scenario.fixture)).toString('base64') });
            return;
        }
        if (path === '/golden/model' && request.method === 'POST') {
            const { step, model, payload } = JSON.parse((await readBody(request)).toString('utf8'));
            const queue = queues.get(step);
            const served = queue ? queue.next(payload) : { kind: 'unknown', error: `model call outside a step (${step})`,
                response: { candidates: [{ content: { role: 'model', parts: [{ text: '{}' }] }, finishReason: 'STOP' }] } };
            appendFileSync(modelLog, `${JSON.stringify({ step, model, kind: served.kind, error: served.error || null,
                served: served.response.candidates[0].content.parts[0], payload })}\n`);
            if (served.error) console.log(`  REPLAY ${step}: ${served.error}`);
            json(served.response);
            return;
        }
        if (path === '/golden/model-log' && request.method === 'POST') {
            appendFileSync(modelLog, `${(await readBody(request)).toString('utf8')}\n`);
            response.end('ok');
            return;
        }
        const stepRoute = path.match(/^\/golden\/step\/(\d+)\/(docx|meta)$/);
        if (stepRoute && request.method === 'POST') {
            const index = Number(stepRoute[1]);
            const body = await readBody(request);
            if (stepRoute[2] === 'docx') pendingDocx.set(index, new Uint8Array(body));
            else await scoreCompletedStep(index, JSON.parse(body.toString('utf8')));
            response.end('ok');
            return;
        }
        if (path === '/golden/done' && request.method === 'POST') {
            const summary = JSON.parse((await readBody(request)).toString('utf8'));
            response.end('ok');
            finish(summary);
            return;
        }

        const file = resolve(distDirectory, `.${path}`);
        if (!file.startsWith(distDirectory) || !existsSync(file)) { response.writeHead(404).end(); return; }
        const types = { '.js': 'application/javascript', '.html': 'text/html', '.css': 'text/css', '.json': 'application/json', '.png': 'image/png', '.map': 'application/json' };
        response.setHeader('Content-Type', types[extname(file)] || 'application/octet-stream');
        if (path === '/taskpane.html') response.end(readFileSync(file, 'utf8').replace('<head>', `<head>${sessionScript}`));
        else response.end(readFileSync(file));
    } catch (error) {
        console.error(`Golden server error: ${error.message}`);
        if (!response.headersSent) response.writeHead(500);
        response.end('Golden server failed');
    }
}

// Word's WebView resolves localhost to ::1 first; serve both loopback addresses only.
const servers = ['127.0.0.1', '::1'].map(host => createServer(tls, handle).listen(PORT, host));
const server = { close(callback) { let open = servers.length; for (const item of servers) item.close(() => { if (--open === 0) callback?.(); }); } };

servers[0].on('listening', async () => {
    console.log(`Golden scenario server ready: https://localhost:${PORT}/taskpane.html (${mode} mode)`);
    console.log(`Artifacts: ${artifacts}`);
    if (flag('--launch')) {
        try {
            // A running Word reads developer registrations only at startup, so open
            // the sideload document in a separate instance (/x) and leave any
            // existing Word session untouched.
            const devSettings = require('office-addin-dev-settings');
            const { generateSideloadFile } = require('office-addin-dev-settings/lib/sideload.js');
            const { OfficeAddinManifest } = require('office-addin-manifest');
            await devSettings.registerAddIn(manifestPath);
            const sideloadFile = await generateSideloadFile('word', await OfficeAddinManifest.readManifestFile(manifestPath));
            const wordExe = argument('--word-exe', 'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE');
            const word = spawn(wordExe, ['/x', sideloadFile], { detached: true, stdio: 'ignore' });
            word.unref();
            launchedWordPid = word.pid;
            console.log(`Started a separate Word instance (PID ${word.pid}) with the golden add-in.`);
        } catch (error) {
            console.error(`Launch failed: ${error.message}`);
            process.exitCode = 1;
            server.close();
        }
    }
});

const limitMs = steps.length * stepTimeoutMs + 10 * 60 * 1000;
setTimeout(() => {
    console.log('Golden scenario time limit reached.');
    finish({ status: 'aborted', error: 'time limit reached' });
}, limitMs).unref();
