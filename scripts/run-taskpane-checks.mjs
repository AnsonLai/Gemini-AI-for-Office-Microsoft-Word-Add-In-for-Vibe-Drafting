// Taskpane checks lane: automates the manual Word checks M1 to M14 of
// docs/plans/2026-11-07-state-machine-fixes.md in the real taskpane in desktop Word.
//
//   node scripts/run-taskpane-checks.mjs --launch [--only M2,M5] [--artifacts-dir PATH]
//   node scripts/run-taskpane-checks.mjs --cleanup     # remove the sideload registration
//
// A scripted responder inside the page answers every Gemini request (no provider
// calls, only a dummy API key). Serves on port 3001 like the golden scenario, so do
// not run both at once. Uses existing localhost development certificates; never
// installs trust. Closes only the Word instance it started.
import { createServer } from 'node:https';
import { execFileSync, spawn } from 'node:child_process';
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { homedir } from 'node:os';
import { dirname, extname, join, resolve } from 'node:path';

const require = createRequire(import.meta.url);
const root = resolve(import.meta.dirname, '..');
const PORT = 3001;
const ADDIN_ID = '5c1f6d0e-7b52-4f0a-9a7e-2d4c8e1b3a90';
// --fixture/--phrase swap the document (the phrase is the text the scripted edits target); the checks assume the NDA by default.
const FIXTURE = argument('--fixture', 'tests/fixtures/golden/sample-nda.docx');
const PHRASE = argument('--phrase', null);

function argument(name, fallback = null) {
    const index = process.argv.indexOf(name);
    if (index < 0) return fallback;
    const value = process.argv[index + 1];
    if (!value || value.startsWith('--')) throw new Error(`Missing value for ${name}`);
    return value;
}
const flag = name => process.argv.includes(name);

const stamp = new Date().toISOString().replace(/[:.]/g, '-');
const artifacts = resolve(root, argument('--artifacts-dir', `.cache/golden/checks-${stamp}`));
const distDirectory = resolve(root, argument('--dist-dir', '.cache/golden/checks-dist'));
const manifestPath = resolve(root, '.cache/golden/checks-manifest.xml');
const only = argument('--only', null)?.split(',').map(id => id.trim().toUpperCase()) || null;

if (flag('--cleanup')) {
    if (existsSync(manifestPath)) await require('office-addin-dev-settings').unregisterAddIn(manifestPath);
    console.log('Taskpane checks add-in registration removed.');
    process.exit(0);
}

mkdirSync(artifacts, { recursive: true });
mkdirSync(dirname(manifestPath), { recursive: true });

if (!flag('--skip-build')) {
    console.log('Building taskpane checks bundle...');
    execFileSync(process.execPath, [require.resolve('webpack/bin/webpack.js'), '--mode', 'development',
        '--env', 'TASKPANE_CHECKS=1', '--output-path', distDirectory], { cwd: root, stdio: 'inherit' });
}

writeFileSync(manifestPath, readFileSync(join(root, 'manifest.xml'), 'utf8')
    .replaceAll('3dcfdb34-70c3-4bfe-8d5e-85089afcf673', ADDIN_ID)
    .replaceAll('Gemini AI for Office', 'Taskpane checks (local)')
    .replace(/(id="TaskpaneButton\.Label" DefaultValue=")Assistant"/g, '$1Checks"')
    .replaceAll('https://localhost:3000/', `https://localhost:${PORT}/`));

const sessionScript = `<script>(function(){window.__CHECKS_SESSION__=${JSON.stringify({ only, phrase: PHRASE })};})();</script>`;
const results = [];
let claimed = false;
let launchedWordPid = null;

async function readBody(request) {
    const parts = [];
    for await (const part of request) parts.push(part);
    return Buffer.concat(parts);
}

function finish(summary) {
    const counts = {};
    for (const result of results) counts[result.status] = (counts[result.status] || 0) + 1;
    const failed = (counts.FAIL || 0) > 0 || summary.status !== 'completed';
    console.log('\n=== Taskpane checks ===');
    for (const result of results) console.log(`${result.id.padEnd(4)} ${result.status.padEnd(7)} ${result.title}`);
    console.log(`\n${JSON.stringify(counts)}${summary.status !== 'completed' ? ` ABORTED at ${summary.phase}: ${summary.error}` : ''}`);
    writeFileSync(join(artifacts, 'report.json'), JSON.stringify({ finishedAt: new Date().toISOString(), summary, counts, results }, null, 2));
    const rows = results.map(item => `| ${item.id} | ${item.status} | ${item.title} | ${item.evidence.replace(/\|/g, '\\|').replace(/\n/g, ' ')} |`);
    writeFileSync(join(artifacts, 'report.md'), `# Taskpane checks\n\n${JSON.stringify(counts)}\n\n| Check | Result | Title | Evidence |\n| --- | --- | --- | --- |\n${rows.join('\n')}\n`);
    console.log(`Artifacts: ${artifacts}`);
    process.exitCode = failed ? 1 : 0;
    // Close only the disposable Word this run started (its document is a temp file).
    if (launchedWordPid && !flag('--keep-word')) {
        try { process.kill(launchedWordPid); } catch { /* already closed */ }
    }
    const unregister = launchedWordPid && !flag('--keep-word')
        ? require('office-addin-dev-settings').unregisterAddIn(manifestPath).catch(error => console.error(`Unregister failed: ${error.message}`))
        : Promise.resolve();
    unregister.finally(() => setTimeout(() => server.close(() => process.exit()), 500));
}

const certs = join(homedir(), '.office-addin-dev-certs');
const tls = { cert: readFileSync(join(certs, 'localhost.crt')), key: readFileSync(join(certs, 'localhost.key')) };
async function handle(request, response) {
    try {
        const path = new URL(request.url, `https://localhost:${PORT}`).pathname;
        response.setHeader('Cache-Control', 'no-store');
        if (path === '/checks/claim' && request.method === 'POST') {
            await readBody(request);
            if (claimed) { response.writeHead(409).end(); return; }
            claimed = true;
            console.log('Run claimed by Word.');
            response.setHeader('Content-Type', 'application/json');
            response.end(JSON.stringify({ fixture: readFileSync(resolve(root, FIXTURE)).toString('base64') }));
            return;
        }
        if (path === '/checks/result' && request.method === 'POST') {
            const result = JSON.parse((await readBody(request)).toString('utf8'));
            results.push(result);
            console.log(`${result.id.padEnd(4)} ${result.status.padEnd(7)} ${result.title} (${Math.round(result.ms / 1000)}s)\n       ${result.evidence}`);
            response.end('ok');
            return;
        }
        if (path === '/checks/done' && request.method === 'POST') {
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
        console.error(`Checks server error: ${error.message}`);
        if (!response.headersSent) response.writeHead(500);
        response.end('Checks server failed');
    }
}

// Word's WebView resolves localhost to ::1 first; serve both loopback addresses only.
const servers = ['127.0.0.1', '::1'].map(host => createServer(tls, handle).listen(PORT, host));
const server = {
    close(callback) {
        let open = servers.length;
        for (const item of servers) item.close(() => { if (--open === 0) callback?.(); });
        for (const item of servers) item.closeAllConnections?.();
    }
};

servers[0].on('listening', async () => {
    console.log(`Taskpane checks server ready: https://localhost:${PORT}/taskpane.html`);
    if (flag('--launch')) {
        try {
            const devSettings = require('office-addin-dev-settings');
            const { generateSideloadFile } = require('office-addin-dev-settings/lib/sideload.js');
            const { OfficeAddinManifest } = require('office-addin-manifest');
            await devSettings.registerAddIn(manifestPath);
            const sharedSideloadFile = await generateSideloadFile('word', await OfficeAddinManifest.readManifestFile(manifestPath));
            // A copy that belongs to this run, so a leftover Word holding the shared name cannot block it.
            const sideloadFile = join(artifacts, `sideload-${Date.now()}.docx`);
            copyFileSync(sharedSideloadFile, sideloadFile);
            const wordExe = argument('--word-exe', 'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE');
            const word = spawn(wordExe, ['/x', sideloadFile], { detached: true, stdio: 'ignore' });
            word.unref();
            launchedWordPid = word.pid;
            console.log(`Started a separate Word instance (PID ${word.pid}) with the checks add-in.`);
        } catch (error) {
            console.error(`Launch failed: ${error.message}`);
            process.exitCode = 1;
            server.close();
        }
    }
});

setTimeout(() => {
    console.log('Taskpane checks time limit reached.');
    finish({ status: 'aborted', error: 'time limit reached', phase: 'time limit' });
}, 20 * 60 * 1000).unref();
