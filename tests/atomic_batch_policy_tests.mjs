import assert from 'assert';
import { readFileSync, readdirSync, statSync } from 'fs';
import { dirname, join, resolve } from 'path';
import { fileURLToPath } from 'url';

// Plugin policy: every batch sent to the library is one transaction, so a
// failed operation leaves the document untouched.
const SRC = resolve(dirname(fileURLToPath(import.meta.url)), '..', 'src');

function walkJavaScriptFiles(root) {
    const files = [];
    for (const entry of readdirSync(root)) {
        const fullPath = join(root, entry);
        if (statSync(fullPath).isDirectory()) files.push(...walkJavaScriptFiles(fullPath));
        else if (/\.(?:js|mjs)$/.test(entry)) files.push(fullPath);
    }
    return files;
}

let callers = 0;
for (const file of walkJavaScriptFiles(SRC)) {
    const source = readFileSync(file, 'utf8');
    if (!source.includes('applyOperationsToDocumentXml')) continue;
    callers++;
    assert.match(
        source,
        /atomic\s*:\s*true/,
        `${file} uses applyOperationsToDocumentXml without an explicit atomic: true transaction policy`
    );
}
assert.ok(callers > 0, 'the scan found no applyOperationsToDocumentXml caller; update this test if the library entry point changed');

console.log(`PASS: ${callers} production file(s) call applyOperationsToDocumentXml with atomic: true`);
