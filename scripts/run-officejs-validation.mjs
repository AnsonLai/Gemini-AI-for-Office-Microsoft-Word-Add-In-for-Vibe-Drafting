// Local-only evidence collector. Uses existing certificates; never installs trust.
import { createServer } from 'node:https';
import { readFileSync, writeFileSync, mkdirSync, existsSync, copyFileSync } from 'node:fs';
import { resolve, join, extname, dirname } from 'node:path';
import { homedir } from 'node:os';
import { execFileSync, spawn } from 'node:child_process';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const root = resolve(import.meta.dirname, '..');
function argument(name) {
    const index = process.argv.indexOf(name);
    if (index < 0) return null;
    if (!process.argv[index + 1] || process.argv[index + 1].startsWith('--')) throw new Error(`Missing value for ${name}`);
    return process.argv[index + 1];
}
const directory = resolve(root, argument('--artifacts-dir') || '.cache/reliability/officejs');
const fixtureManifest = argument('--fixture-manifest');
const startupProfile = process.argv.includes('--startup');
const distDirectory = resolve(root, argument('--dist-dir') || 'dist');
if (process.argv.includes('--cleanup')) {
    await require('office-addin-dev-settings').unregisterAddIn(join(directory, 'officejs-manifest.xml'));
    console.log('Local validation add-in registration removed.');
    process.exit(0);
}
mkdirSync(directory, { recursive: true });
const startedAt = new Date().toISOString();
writeFileSync(join(directory, startupProfile ? 'startup-report.json' : 'officejs-report.json'), JSON.stringify({ status: 'pending', startedAt }));
writeFileSync(join(directory, 'officejs-fixtures.json'), JSON.stringify({ cases: [], controls: [] }));
if (!fixtureManifest && !startupProfile) execFileSync(process.execPath, [
    join(root, 'tests/direct_document_fidelity_manifest_tests.mjs'), '--export-dir', directory
], { stdio: 'inherit' });
const inputManifest = fixtureManifest ? resolve(root, fixtureManifest) : join(directory, 'manifest.json');
const sourceDirectory = dirname(inputManifest);
const manifest = startupProfile ? { cases: [] } : JSON.parse(readFileSync(inputManifest, 'utf8'));
const cases = manifest.cases;
if (!startupProfile && !cases.length) throw new Error('Expected runnable direct document-tool fixtures');
for (const testCase of cases) {
    if (!/^[a-z0-9-]+$/.test(testCase.name)) throw new Error(`Invalid fixture case name: ${testCase.name}`);
    copyFileSync(resolve(sourceDirectory, testCase.source), join(directory, `${testCase.name}-source.docx`));
}
const certs = join(homedir(), '.office-addin-dev-certs');
const cert = readFileSync(join(certs, 'localhost.crt'));
const key = readFileSync(join(certs, 'localhost.key'));
const port = Number(argument('--port') || '3002');
const id = startupProfile ? 'a81aa9d8-d01b-4196-b79d-e4b92030fc83' : 'a81aa9d8-d01b-4196-b79d-e4b92030fc82';
const manifestPath = join(directory, 'officejs-manifest.xml');
writeFileSync(manifestPath, readFileSync(join(root, 'manifest.xml'), 'utf8')
    .replaceAll('3dcfdb34-70c3-4bfe-8d5e-85089afcf673', id)
    .replaceAll('Gemini AI for Office', 'Local Word transport validation')
    .replaceAll('https://localhost:3000/taskpane.html', `https://localhost:${port}/${startupProfile ? 'taskpane' : 'officejs-validation'}.html`));
let claimed = false;
let launchedWordPid = null;
let addInRegistered = false;
let shuttingDown = false;
let timeLimit;
const requestHandler = async (request, response) => {
    try {
        const path = new URL(request.url, `https://localhost:${port}`).pathname;
        const workflowArtifact = path.match(/^\/workflow-artifact\/([a-z0-9-]+)\/([a-z0-9-]+)$/);
        response.setHeader('Cache-Control', 'no-store');
        if (startupProfile && path === '/startup-result' && request.method === 'POST') {
            const parts = []; let size = 0;
            for await (const part of request) { size += part.length; if (size > 1024 * 1024) throw new Error('Startup report too large'); parts.push(part); }
            const result = JSON.parse(Buffer.concat(parts).toString('utf8'));
            writeFileSync(join(directory, 'startup-report.json'), JSON.stringify({ ...result, startedAt, checkedAt: new Date().toISOString() }, null, 2));
            console.log(`Taskpane startup result: ${result.status}`);
            response.end('ok');
            setImmediate(() => { void finishRunner(); });
            return;
        }
        if (path === '/claim' && request.method === 'POST') {
            if (claimed) { response.writeHead(409).end(); return; }
            claimed = true;
            response.setHeader('Content-Type', 'application/json');
            response.end(JSON.stringify(cases)); return;
        }
        const testCase = cases.find(c => path.endsWith('/' + c.name) || workflowArtifact?.[1] === c.name);
        if (path.startsWith('/source/') && testCase) {
            response.end(readFileSync(resolve(sourceDirectory, testCase.source)).toString('base64')); return;
        }
        if (((path.startsWith('/artifact/') || path.startsWith('/prepared-source/') || path.startsWith('/checkpoint/') || path.startsWith('/restored/')
            || workflowArtifact || path.startsWith('/restored-source/') || path.startsWith('/restore-diagnostics/')
            || path.startsWith('/revision-diagnostics/')) && testCase || path === '/result') && request.method === 'POST') {
            const parts = []; let size = 0;
            for await (const part of request) { size += part.length; if (size > 20 * 1024 * 1024) throw new Error('Artifact too large'); parts.push(part); }
            const payload = Buffer.concat(parts);
            if (path === '/result') {
                const result = JSON.parse(payload.toString('utf8'));
                const restorationControls = Array.isArray(result.restorationControls) ? result.restorationControls : [];
                const requiredRestorationCases = [
                    'same-level-insert-redlines-disabled-restores-tracking',
                    'direct-untracked-replacement-preserves-pending-revision',
                    'direct-threaded-comments-sequential-merge-and-restore'
                ].filter(name => cases.some(testCase => testCase.name === name));
                if (result.status === 'passed' && cases.some(c => c.usesPreparedSource && !existsSync(join(directory, `${c.name}-prepared-source.docx`)))) {
                    throw new Error('Passing host result is missing a prepared source package');
                }
                if (result.status === 'passed' && cases.some(c => !existsSync(join(directory, `${c.name}-officejs.docx`)))) {
                    throw new Error('Passing host result is missing an exported Word document');
                }
                if (result.status === 'passed' && cases.some(c => !c.expectedPrewriteRefusal
                    && !existsSync(join(directory, `${c.name}-checkpoint.xml`)))) {
                    throw new Error('Passing host result is missing a source checkpoint');
                }
                if (result.status === 'passed' && (restorationControls.length !== requiredRestorationCases.length
                    || requiredRestorationCases.some(name => !restorationControls.some(control => control.case === name))
                    || restorationControls.some(control => !cases.some(c => c.name === control.case)
                        || !existsSync(join(directory, `${control.case}-restored.docx`))))) {
                    throw new Error('Passing host result is missing a required IndexedDB checkpoint-restoration export');
                }
                writeFileSync(join(directory, 'officejs-report.json'), JSON.stringify({ ...result, startedAt, checkedAt: new Date().toISOString() }, null, 2));
                if (result.status === 'passed') {
                    const oracleCases = cases.filter(c => !c.expectedPrewriteRefusal).map(c => ({
                        ...c,
                        ...(result.fixtureExpectations?.[c.name] || {}),
                        source: c.usesPreparedSource ? join(directory, `${c.name}-prepared-source.docx`) : join(directory, `${c.name}-source.docx`),
                        checkpoint: join(directory, `${c.name}-checkpoint.xml`),
                        ...(c.accepted ? { accepted: resolve(sourceDirectory, c.accepted) } : {}),
                        ...(c.rejected ? { rejected: resolve(sourceDirectory, c.rejected) } : {}),
                        tracked: `${c.name}-officejs.docx` }));
                    for (const control of restorationControls) {
                        const restoredCase = oracleCases.find(c => c.name === control.case);
                        const restoredArtifact = `${restoredCase.name}-restored.docx`;
                        const originalNumbering = restoredCase.expectedNumbering?.source;
                        oracleCases.push({
                            ...restoredCase,
                            name: `${restoredCase.name}-checkpoint-restoration`,
                            tracked: restoredArtifact,
                            accepted: restoredArtifact,
                            rejected: restoredArtifact,
                            expectedMinimumBodyRevisions: Number.isInteger(restoredCase.expectedRestoredBodyRevisions)
                                ? undefined : control.checkpointRevisionElementCount > 0 ? 1 : undefined,
                            expectedExactBodyRevisions: Number.isInteger(restoredCase.expectedRestoredBodyRevisions)
                                ? restoredCase.expectedRestoredBodyRevisions
                                : control.checkpointRevisionElementCount === 0 ? 0 : undefined,
                            expectedSourceText: restoredCase.expectedSourceText,
                            expectedAcceptedText: restoredCase.expectedRestoredAcceptedText || restoredCase.expectedSourceText,
                            expectedRejectedText: restoredCase.expectedRestoredRejectedText || restoredCase.expectedSourceText,
                            expectedComments: Number.isInteger(control.checkpointCommentCount)
                                ? control.checkpointCommentCount : restoredCase.expectedComments ?? 0,
                            ...(Array.isArray(restoredCase.expectedRestoredThreadIdentities)
                                ? { expectedThreadIdentities: restoredCase.expectedRestoredThreadIdentities }
                                : {}),
                            ...(originalNumbering ? { expectedNumbering: {
                                source: originalNumbering, accepted: originalNumbering, rejected: originalNumbering
                            } } : {})
                        });
                    }
                    writeFileSync(join(directory, 'officejs-fixtures.json'), JSON.stringify({ cases: oracleCases, controls: [] }, null, 2));
                }
                console.log(`Office.js result: ${result.status}`);
                response.end('ok');
                setImmediate(() => { void finishRunner(); });
                return;
            } else {
                const artifact = workflowArtifact ? `${testCase.name}-${workflowArtifact[2]}-officejs.docx`
                    : path.startsWith('/checkpoint/') ? `${testCase.name}-checkpoint.xml`
                    : path.startsWith('/restore-diagnostics/') ? `${testCase.name}-restore-diagnostics.json`
                        : path.startsWith('/revision-diagnostics/') ? `${testCase.name}-revision-diagnostics.json`
                        : path.startsWith('/restored-source/') ? `${testCase.name}-restored-source.xml`
                            : `${testCase.name}-${path.startsWith('/prepared-source/') ? 'prepared-source' : path.startsWith('/restored/') ? 'restored' : 'officejs'}.docx`;
                writeFileSync(join(directory, artifact), payload);
            }
            response.end('ok'); return;
        }
        const file = resolve(distDirectory, '.' + path);
        if (!file.startsWith(distDirectory + '/') && !file.startsWith(distDirectory + '\\')) { response.writeHead(404).end(); return; }
        if (!existsSync(file)) { response.writeHead(404).end(); return; }
        response.setHeader('Content-Type', extname(file) === '.js' ? 'application/javascript' : extname(file) === '.html' ? 'text/html' : 'application/octet-stream');
        if (startupProfile && path === '/taskpane.html') {
            const diagnostics = `<script>window.addEventListener('error',function(event){fetch('/startup-result',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({status:'failed',stage:'window-error',message:String(event.message||'Startup error'),line:event.lineno,column:event.colno})}).catch(function(){});},true);window.addEventListener('unhandledrejection',function(event){fetch('/startup-result',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({status:'failed',stage:'unhandled-rejection',message:String(event.reason&&event.reason.message||'Startup rejection')})}).catch(function(){});});</script>`;
            response.end(readFileSync(file, 'utf8').replace('<head>', '<head>' + diagnostics));
        } else response.end(readFileSync(file));
    } catch (error) { console.error(error.message); response.writeHead(500).end('Local collector failed'); }
};
const listeners = ['127.0.0.1', '::1'].map(host => createServer({ cert, key }, requestHandler).listen(port, host));
const server = { close(callback) {
    const openListeners = listeners.filter(listener => listener.listening);
    let remaining = openListeners.length;
    if (!remaining) { callback?.(); return; }
    for (const listener of openListeners) {
        listener.close(() => { if (--remaining === 0) callback?.(); });
        // Word's WebView can retain HTTP keep-alive sockets after the final
        // evidence post; shutdown owns these connections, so close them too.
        listener.closeAllConnections?.();
    }
} };
async function finishRunner({ failure = null } = {}) {
    if (shuttingDown) return;
    shuttingDown = true;
    clearTimeout(timeLimit);
    if (failure) writeFileSync(join(directory, startupProfile ? 'startup-report.json' : 'officejs-report.json'), JSON.stringify({
        status: 'failed', error: failure, startedAt, checkedAt: new Date().toISOString()
    }, null, 2));
    if (launchedWordPid) {
        try { process.kill(launchedWordPid); } catch { /* the isolated Word process already exited */ }
    }
    if (addInRegistered) {
        try { await require('office-addin-dev-settings').unregisterAddIn(manifestPath); }
        catch (error) { console.error(`Validation add-in unregister failed: ${error.message}`); }
        addInRegistered = false;
    }
    server.close();
}

listeners.forEach(listener => listener.on('error', error => {
    console.error(`Validation server bind failed: ${error.message}`);
    process.exitCode = 1;
    void finishRunner({ failure: error.message });
}));
listeners[0].on('listening', async () => {
    console.log(`Local validation collector ready: https://localhost:${port}/${startupProfile ? 'taskpane' : 'officejs-validation'}.html`);
    if (process.argv.includes('--launch')) {
        try {
            // Read the developer registration at startup in a separate Word
            // instance. This leaves any existing user Word session untouched.
            const settings = require('office-addin-dev-settings');
            const { generateSideloadFile } = require('office-addin-dev-settings/lib/sideload.js');
            const { OfficeAddinManifest } = require('office-addin-manifest');
            await settings.registerAddIn(manifestPath);
            addInRegistered = true;
            const sideloadFile = await generateSideloadFile('word', await OfficeAddinManifest.readManifestFile(manifestPath));
            const wordExe = argument('--word-exe') || 'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE';
            const word = spawn(wordExe, ['/x', sideloadFile], { detached: true, stdio: 'ignore' });
            word.unref();
            launchedWordPid = word.pid;
            word.once('error', error => {
                console.error(`Validation Word launch failed: ${error.message}`);
                process.exitCode = 1;
                void finishRunner({ failure: error.message });
            });
            console.log(`Started a separate Word instance (PID ${word.pid}) with the validation add-in.`);
        } catch (error) {
            console.error(`Validation launch failed: ${error.message}`);
            process.exitCode = 1;
            await finishRunner({ failure: error.message });
        }
    }
});
// Bound the collector lifetime; evidence stays available on timeout.
timeLimit = setTimeout(() => {
    console.log('Collector time limit reached.');
    void finishRunner({ failure: 'Local validation collector timed out.' });
}, 10 * 60 * 1000).unref();
