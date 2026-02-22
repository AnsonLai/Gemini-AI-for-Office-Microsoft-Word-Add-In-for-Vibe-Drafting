import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const reconciliationDir = path.join(__dirname, '../src/taskpane/modules/reconciliation');
const integrationDir = path.join(reconciliationDir, 'integration');
const standalonePath = path.join(reconciliationDir, 'standalone.js');
const indexPath = path.join(reconciliationDir, 'index.js');
const wordAddinEntryPath = path.join(reconciliationDir, 'word-addin-entry.js');
const allowedExternalImports = new Set(['diff-match-patch']);

const forbiddenPatterns = [
    /Office\./,
    /Word\./,
    /context\.sync/,
    /paragraph\.(getOoxml|insertOoxml)/
];

function stripComments(source) {
    return source
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/\/\/.*$/gm, '');
}

function isPathWithin(parentDir, candidatePath) {
    const rel = path.relative(parentDir, candidatePath);
    return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
}

async function collectJsFilesRecursively(dirPath) {
    const entries = await fs.readdir(dirPath, { withFileTypes: true });
    const files = [];

    for (const entry of entries) {
        const fullPath = path.join(dirPath, entry.name);
        if (entry.isDirectory()) {
            if (entry.name === 'integration') continue;
            const nested = await collectJsFilesRecursively(fullPath);
            files.push(...nested);
            continue;
        }

        if (entry.isFile() && entry.name.endsWith('.js')) {
            files.push(fullPath);
        }
    }

    return files;
}

function extractImportSpecifiers(sourceNoComments) {
    const specifiers = new Set();
    const staticImportRe = /\bimport\s+[^'"]*?\s+from\s+['"]([^'"]+)['"]/gms;
    const sideEffectImportRe = /\bimport\s+['"]([^'"]+)['"]/gms;
    const exportFromRe = /\bexport\s+[^'"]*?\s+from\s+['"]([^'"]+)['"]/gms;
    const dynamicImportRe = /\bimport\s*\(\s*['"]([^'"]+)['"]\s*\)/gms;

    let match = null;
    while ((match = staticImportRe.exec(sourceNoComments)) !== null) {
        specifiers.add(match[1]);
    }
    while ((match = sideEffectImportRe.exec(sourceNoComments)) !== null) {
        specifiers.add(match[1]);
    }
    while ((match = exportFromRe.exec(sourceNoComments)) !== null) {
        specifiers.add(match[1]);
    }
    while ((match = dynamicImportRe.exec(sourceNoComments)) !== null) {
        specifiers.add(match[1]);
    }

    return [...specifiers];
}

function resolveImportPath(importerPath, specifier) {
    if (specifier.startsWith('.')) {
        return path.resolve(path.dirname(importerPath), specifier);
    }

    if (specifier.startsWith('/')) {
        return path.resolve(specifier);
    }

    return null;
}

async function run() {
    const standaloneSource = stripComments(await fs.readFile(standalonePath, 'utf8'));
    const standaloneImports = extractImportSpecifiers(standaloneSource);
    for (const specifier of standaloneImports) {
        const resolved = resolveImportPath(standalonePath, specifier);
        if (resolved && isPathWithin(integrationDir, resolved)) {
            throw new Error(`standalone.js must not import from integration/: ${specifier}`);
        }
    }

    const files = await collectJsFilesRecursively(reconciliationDir);
    const filesToCheck = files.filter(filePath =>
        filePath !== indexPath &&
        filePath !== wordAddinEntryPath &&
        !isPathWithin(integrationDir, filePath)
    );

    for (const fullPath of filesToCheck) {
        const sourceNoComments = stripComments(await fs.readFile(fullPath, 'utf8'));
        const relativePath = path.relative(reconciliationDir, fullPath);
        for (const pattern of forbiddenPatterns) {
            if (pattern.test(sourceNoComments)) {
                throw new Error(`Forbidden Word API pattern ${pattern} found in ${relativePath}`);
            }
        }

        const imports = extractImportSpecifiers(sourceNoComments);
        for (const specifier of imports) {
            const resolved = resolveImportPath(fullPath, specifier);

            if (resolved && isPathWithin(integrationDir, resolved)) {
                throw new Error(`${relativePath} imports integration-only module: ${specifier}`);
            }

            if (resolved) {
                if (!isPathWithin(reconciliationDir, resolved)) {
                    throw new Error(`${relativePath} imports outside reconciliation/: ${specifier}`);
                }
                continue;
            }

            if (!allowedExternalImports.has(specifier)) {
                throw new Error(`${relativePath} imports disallowed external package: ${specifier}`);
            }
        }
    }

    console.log('PASS: reconciliation core has no Word API markers and no invalid cross-boundary imports');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
