import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const reconciliationDir = path.join(__dirname, '../src/taskpane/modules/reconciliation');
const standalonePath = path.join(reconciliationDir, 'standalone.js');

const forbiddenPatterns = [
    /Word\.run/g,
    /Office\./g,
    /context\.sync/g,
    /paragraph\.(getOoxml|insertOoxml)/g
];

async function run() {
    const standaloneSource = await fs.readFile(standalonePath, 'utf8');
    if (standaloneSource.includes("./integration.js") || standaloneSource.includes('./integration.js')) {
        throw new Error('standalone.js must not import integration.js');
    }

    const files = await fs.readdir(reconciliationDir);
    for (const file of files) {
        if (!file.endsWith('.js')) continue;
        if (file === 'integration.js') continue;

        const fullPath = path.join(reconciliationDir, file);
        const source = await fs.readFile(fullPath, 'utf8');
        const sourceNoComments = source
            .replace(/\/\*[\s\S]*?\*\//g, '')
            .replace(/\/\/.*$/gm, '');

        for (const pattern of forbiddenPatterns) {
            if (pattern.test(sourceNoComments)) {
                throw new Error(`Forbidden Word API pattern ${pattern} found in ${file}`);
            }
        }
    }

    console.log('PASS: standalone graph has no direct Word API markers outside integration.js');
}

run().catch(err => {
    console.error('FAIL:', err.message);
    process.exit(1);
});
