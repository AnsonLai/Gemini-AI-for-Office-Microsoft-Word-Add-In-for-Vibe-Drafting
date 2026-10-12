import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { JSDOM } from 'jsdom';
import { populateModelSetting, readModelSetting } from '../src/taskpane/modules/config/model-settings.js';
const html = readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf8');
const dom = new JSDOM(html);
globalThis.document = dom.window.document;
for (const type of ['fast', 'slow']) {
  const select = document.getElementById(`model-select-${type}`);
  const custom = document.getElementById(`model-custom-${type}`);
  populateModelSetting(type, 'gemini-2.5-pro');
  assert.equal(readModelSetting(type), 'gemini-2.5-pro');
  assert.equal(custom.hidden, true);
  populateModelSetting(type, 'gemini-example-001');
  assert.equal(select.value, 'custom');
  assert.equal(custom.hidden, false);
  assert.equal(readModelSetting(type), 'gemini-example-001');
  custom.value = '  gemini-other-preview  ';
  assert.equal(readModelSetting(type), 'gemini-other-preview');
  for (const invalid of ['', ' ', '../model', 'models/model', 'model?key=secret']) {
    custom.value = invalid;
    assert.throws(() => readModelSetting(type), /model ID/);
  }
  select.value = 'gemini-2.5-flash';
  select.dispatchEvent(new dom.window.Event('change'));
  assert.equal(custom.hidden, true);
  assert.equal(readModelSetting(type), 'gemini-2.5-flash');
}
populateModelSetting('fast', 'gemini-flash-latest');
assert.equal(readModelSetting('fast'), 'gemini-flash-latest');
assert.match(document.querySelector('#model-select-fast option[selected]').textContent, /Default/);
console.log('PASS custom model settings round trips and validation');
