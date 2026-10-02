// Fixed physical-line surface used by the 2026-10-02 reliability milestones.
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { resolve } from 'node:path';
const root = resolve(import.meta.dirname, '..');
const roots = [
  'src/taskpane/taskpane.js', 'src/taskpane/modules/commands',
  'src/taskpane/modules/docx-redline-js-integration',
  'src/taskpane/modules/chat', 'src/commands/commands.js',
];
const refIndex = process.argv.indexOf('--ref');
const ref = refIndex < 0 ? null : process.argv[refIndex + 1];
if (refIndex >= 0 && (!ref || ref.startsWith('--'))) throw new Error('--ref requires a Git revision.');
const git = args => execFileSync('git', args, { cwd: root, encoding: 'utf8' });
const walk = path => !existsSync(resolve(root, path)) ? [] : path.endsWith('.js') ? [path]
  : readdirSync(resolve(root, path), { withFileTypes: true }).flatMap(entry => entry.isDirectory()
    ? walk(`${path}/${entry.name}`) : entry.name.endsWith('.js') ? [`${path}/${entry.name}`] : []);
const files = ref ? git(['ls-tree', '-r', '--name-only', ref, '--', ...roots]).trim().split(/\r?\n/).filter(path => path.endsWith('.js'))
  : roots.flatMap(walk);
const groups = roots.map(path => ({ path, files: 0, lines: 0 }));
for (const path of files) {
  const source = ref ? git(['show', `${ref}:${path}`]) : readFileSync(resolve(root, path), 'utf8');
  const lines = (source.match(/\r\n|\r|\n/g) || []).length + (source.length && !/[\r\n]$/.test(source) ? 1 : 0);
  const group = groups.find(group => path === group.path || path.startsWith(`${group.path}/`));
  group.lines += lines; group.files++;
}
console.log(JSON.stringify({ ref: ref || 'working-tree', groups,
  lines: groups.reduce((total, group) => total + group.lines, 0), files: files.length }, null, 2));
