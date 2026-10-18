import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

// Phase-2 WP4: glance checks read library inspection text, never Word's body.text,
// and the retired expedient files must not return.
const taskpane = readFileSync(resolve(root, 'src/taskpane/taskpane.js'), 'utf8');
assert.doesNotMatch(taskpane, /\bbody\s*\.\s*text\b/, 'taskpane.js must not read body.text');
assert.doesNotMatch(taskpane, /\bbody\s*\.\s*load\s*\(\s*['"`][^'"`]*\btext\b/, 'taskpane.js must not load the body text property');
assert.doesNotMatch(taskpane, /getRange\(\s*['"`]Whole['"`]\s*\)\s*\.\s*text\b/, 'taskpane.js must not read whole-range text');

for (const retired of ['src/taskpane/debug_structure.js', 'src/commands/commands.js']) {
  assert.equal(existsSync(resolve(root, retired)), false, `${retired} was retired in phase-2 WP4 and must stay absent`);
}

console.log('PASS: phase-2 WP4 dead code stays removed and taskpane.js does not read body.text');
