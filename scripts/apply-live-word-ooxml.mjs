// Transport adapter for the live Word oracle. Calls the production add-in
// bridge against Word's own exported scope XML; never reconstructs a package.
import { readFileSync, writeFileSync } from 'node:fs';
import { executePureOoxmlBatch } from '../src/taskpane/modules/docx-redline-js-integration/word-operation-runner.js';

function argument(name) {
    const index = process.argv.indexOf(name);
    if (index < 0 || !process.argv[index + 1] || process.argv[index + 1].startsWith('--')) throw new Error(`${name} requires a path`);
    return process.argv[index + 1];
}
const input = readFileSync(argument('--input'), 'utf8').replace(/^\uFEFF/, '');
const operations = JSON.parse(readFileSync(argument('--operations'), 'utf8').replace(/^\uFEFF/, ''));
const writes = [];
const scope = { getOoxml: () => ({ value: input }), insertOoxml: value => writes.push(value) };
const result = await executePureOoxmlBatch({ sync: async () => {} }, scope, operations, { author: 'WP6 reviewer' });
if (result.status !== 'ok' || !result.written || writes.length !== 1) {
    throw new Error(`Live bridge did not produce one successful write: ${JSON.stringify(result)}`);
}
writeFileSync(argument('--output'), writes[0], 'utf8');
console.log(`Live Word bridge: ${operations.length} operations, one insertion payload`);
