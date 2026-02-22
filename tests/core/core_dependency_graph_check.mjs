import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const reconciliationDir = path.join(__dirname, '../../src/taskpane/modules/reconciliation');
const integrationDir = path.join(reconciliationDir, 'integration');
const indexPath = path.join(reconciliationDir, 'index.js');
const wordAddinEntryPath = path.join(reconciliationDir, 'word-addin-entry.js');

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

function isNpmPackageImport(specifier) {
    return !specifier.startsWith('.') && !specifier.startsWith('/');
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
    const files = await collectJsFilesRecursively(reconciliationDir);
    const filesToCheck = files.filter(filePath =>
        filePath !== indexPath &&
        filePath !== wordAddinEntryPath &&
        !isPathWithin(integrationDir, filePath)
    );
    const violations = [];

    for (const filePath of filesToCheck) {
        const relativePath = path.relative(reconciliationDir, filePath);
        const sourceNoComments = stripComments(await fs.readFile(filePath, 'utf8'));
        const imports = extractImportSpecifiers(sourceNoComments);

        for (const specifier of imports) {
            if (isNpmPackageImport(specifier)) {
                continue;
            }

            const resolved = resolveImportPath(filePath, specifier);
            if (!resolved || !isPathWithin(reconciliationDir, resolved)) {
                violations.push(`${relativePath} -> ${specifier}`);
            }
        }
    }

    if (violations.length > 0) {
        throw new Error(
            `Core dependency graph violations detected:\n${violations.map(line => ` - ${line}`).join('\n')}`
        );
    }

    console.log('PASS: core dependency graph stays within reconciliation/ or npm package imports');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});

