// Local-only evidence collector. Uses existing certificates; never installs trust.
import { createServer } from 'node:https';
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { resolve, join, extname, dirname } from 'node:path';
import { homedir } from 'node:os';
import { execFileSync } from 'node:child_process';
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
if (!fixtureManifest && !startupProfile) execFileSync(process.execPath, [join(root, 'tests/ooxml_formatting_visual_tests.mjs'), '--export-dir', directory], { stdio: 'inherit' });
const inputManifest = fixtureManifest ? resolve(root, fixtureManifest) : join(directory, 'manifest.json');
const sourceDirectory = dirname(inputManifest);
const manifest = startupProfile ? { cases: [] } : JSON.parse(readFileSync(inputManifest, 'utf8'));
const cases = fixtureManifest ? manifest.cases : manifest.cases.filter(c => ['word-addin-plain-replacement', 'word-addin-sibling-reply'].includes(c.name));
if (!startupProfile && (!cases.length || (!fixtureManifest && cases.length !== 2))) throw new Error('Expected runnable native fixtures');
const certs = join(homedir(), '.office-addin-dev-certs');
const cert = readFileSync(join(certs, 'localhost.crt'));
const key = readFileSync(join(certs, 'localhost.key'));
const port = 3000;
const id = startupProfile ? 'a81aa9d8-d01b-4196-b79d-e4b92030fc83' : 'a81aa9d8-d01b-4196-b79d-e4b92030fc82';
const manifestPath = join(directory, 'officejs-manifest.xml');
writeFileSync(manifestPath, readFileSync(join(root, 'manifest.xml'), 'utf8')
    .replaceAll('3dcfdb34-70c3-4bfe-8d5e-85089afcf673', id)
    .replaceAll('Gemini AI for Office', 'Local Word transport validation')
    .replaceAll('https://localhost:3000/taskpane.html', `https://localhost:3000/${startupProfile ? 'taskpane' : 'officejs-validation'}.html`));
let claimed = false;
const server = createServer({ cert, key }, async (request, response) => {
    try {
        const path = new URL(request.url, `https://localhost:${port}`).pathname;
        response.setHeader('Cache-Control', 'no-store');
        if (startupProfile && path === '/startup-result' && request.method === 'POST') {
            const parts = []; let size = 0;
            for await (const part of request) { size += part.length; if (size > 1024 * 1024) throw new Error('Startup report too large'); parts.push(part); }
            const result = JSON.parse(Buffer.concat(parts).toString('utf8'));
            writeFileSync(join(directory, 'startup-report.json'), JSON.stringify({ ...result, startedAt, checkedAt: new Date().toISOString() }, null, 2));
            console.log(`Taskpane startup result: ${result.status}`);
            response.end('ok'); return;
        }
        if (path === '/claim' && request.method === 'POST') {
            if (claimed) { response.writeHead(409).end(); return; }
            claimed = true;
            response.setHeader('Content-Type', 'application/json');
            response.end(JSON.stringify(cases)); return;
        }
        const testCase = cases.find(c => path.endsWith('/' + c.name));
        if (path.startsWith('/source/') && testCase) {
            response.end(readFileSync(resolve(sourceDirectory, testCase.source)).toString('base64')); return;
        }
        if (((path.startsWith('/artifact/') || path.startsWith('/prepared-source/')) && testCase || path === '/result') && request.method === 'POST') {
            const parts = []; let size = 0;
            for await (const part of request) { size += part.length; if (size > 20 * 1024 * 1024) throw new Error('Artifact too large'); parts.push(part); }
            const payload = Buffer.concat(parts);
            if (path === '/result') {
                const result = JSON.parse(payload.toString('utf8'));
                if (result.status === 'passed' && cases.some(c => c.usesPreparedSource && !existsSync(join(directory, `${c.name}-prepared-source.docx`)))) {
                    throw new Error('Passing host result is missing a prepared source package');
                }
                writeFileSync(join(directory, 'officejs-report.json'), JSON.stringify({ ...result, startedAt, checkedAt: new Date().toISOString() }, null, 2));
                if (result.status === 'passed') {
                    writeFileSync(join(directory, 'officejs-fixtures.json'), JSON.stringify({ cases: cases.map(c => ({ ...c,
                        source: c.usesPreparedSource ? join(directory, `${c.name}-prepared-source.docx`) : resolve(sourceDirectory, c.source),
                        ...(c.accepted ? { accepted: resolve(sourceDirectory, c.accepted) } : {}),
                        ...(c.rejected ? { rejected: resolve(sourceDirectory, c.rejected) } : {}),
                        tracked: `${c.name}-officejs.docx` })), controls: [] }, null, 2));
                }
                console.log(`Office.js result: ${result.status}`);
            } else writeFileSync(join(directory, `${testCase.name}-${path.startsWith('/prepared-source/') ? 'prepared-source' : 'officejs'}.docx`), payload);
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
});
server.listen(port, '127.0.0.1', async () => {
    console.log(`Local validation collector ready: https://localhost:${port}/${startupProfile ? 'taskpane' : 'officejs-validation'}.html`);
    if (process.argv.includes('--launch')) {
        try {
            const settings = require('office-addin-dev-settings');
            await settings.registerAddIn(manifestPath);
            await settings.sideloadAddIn(manifestPath, 'word', true, 'desktop');
        } catch (error) {
            console.error(`Validation launch failed: ${error.message}`);
            process.exitCode = 1;
            server.close();
        }
    }
});
// Bound the collector lifetime; evidence stays available on timeout.
setTimeout(() => { console.log('Collector time limit reached'); server.close(); }, 10 * 60 * 1000).unref();
