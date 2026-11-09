import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { JSDOM } from 'jsdom';
import { saveApiKeySetting, saveModelSettingFrom, saveSystemMessageSetting } from '../src/taskpane/modules/config/settings-save.js';
import { populateModelSetting, readModelSetting } from '../src/taskpane/modules/config/model-settings.js';

const html = readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf8');
const dom = new JSDOM(html);
globalThis.document = dom.window.document;

const makeStorage = (initial = {}) => {
  const data = new Map(Object.entries(initial));
  return {
    getItem: key => (data.has(key) ? data.get(key) : null),
    setItem: (key, value) => { data.set(key, String(value)); },
    removeItem: key => { data.delete(key); },
    has: key => data.has(key),
  };
};

// The API key saves on change; an empty key removes it, so the no-key state can be reached again.
{
  const storage = makeStorage();
  assert.deepEqual(saveApiKeySetting(storage, '  AIza-example  '), { stored: true });
  assert.equal(storage.getItem('geminiApiKey'), 'AIza-example', 'the key is trimmed');
  assert.deepEqual(saveApiKeySetting(storage, ''), { stored: false });
  assert.equal(storage.has('geminiApiKey'), false, 'an empty key removes the stored key');
  saveApiKeySetting(storage, 'AIza-example');
  assert.deepEqual(saveApiKeySetting(storage, '   '), { stored: false });
  assert.equal(storage.has('geminiApiKey'), false, 'a blank key removes it too');
}

// A model saves on its own; a custom model with an invalid id is not saved and reports the error.
{
  const storage = makeStorage({ geminiModelFast: 'gemini-flash-latest', geminiModelSlow: 'gemini-2.5-pro' });
  populateModelSetting('fast', 'gemini-flash-latest');
  populateModelSetting('slow', 'gemini-2.5-pro');
  const fast = document.getElementById('model-select-fast');
  fast.value = 'gemini-2.5-flash';
  assert.deepEqual(saveModelSettingFrom(storage, 'fast', readModelSetting), { ok: true });
  assert.equal(storage.getItem('geminiModelFast'), 'gemini-2.5-flash');
  assert.equal(storage.getItem('geminiModelSlow'), 'gemini-2.5-pro', 'the other model is untouched');

  fast.value = 'custom';
  document.getElementById('model-custom-fast').value = '../model';
  const refused = saveModelSettingFrom(storage, 'fast', readModelSetting);
  assert.equal(refused.ok, false);
  assert.match(refused.error, /model ID/);
  assert.equal(storage.getItem('geminiModelFast'), 'gemini-2.5-flash', 'an invalid id is not saved');

  document.getElementById('model-custom-fast').value = ' gemini-example-001 ';
  assert.deepEqual(saveModelSettingFrom(storage, 'fast', readModelSetting), { ok: true });
  assert.equal(storage.getItem('geminiModelFast'), 'gemini-example-001');
}

// The system prompt saves as typed.
{
  const storage = makeStorage();
  saveSystemMessageSetting(storage, 'Be brief.');
  assert.equal(storage.getItem('geminiSystemMessage'), 'Be brief.');
}

// Settings has no Save button any more: every field saves on change, and Back only leaves.
assert.equal(document.getElementById('save-api-key'), null);
assert.ok(document.getElementById('back-to-main'));
assert.ok(document.getElementById('settings-status'));

console.log('PASS settings save on change: key (including removal), models, system prompt, no Save button');
