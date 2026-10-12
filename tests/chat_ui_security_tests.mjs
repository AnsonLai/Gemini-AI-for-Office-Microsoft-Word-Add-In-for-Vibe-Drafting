import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { JSDOM } from 'jsdom';

const dom = new JSDOM('<div id="chat-messages"></div>');
globalThis.window = dom.window;
globalThis.document = dom.window.document;
const { addMessageToChat, updateSystemMessage } = await import('../src/taskpane/modules/chat/chat-ui.js');

function assertSafe(element) {
  assert.equal(element.querySelector('script,iframe,object,embed,svg,math,style,form,input,img'), null);
  for (const node of element.querySelectorAll('*')) {
    for (const attr of node.attributes) {
      assert.ok(!/^on/i.test(attr.name), `Unexpected handler: ${attr.name}`);
      assert.ok(!/^\s*(javascript|data|vbscript):/i.test(attr.value), `Unsafe URL: ${attr.value}`);
    }
  }
}
const attacks = [
  '<img src=x onerror="window.pwned=1">',
  '<script>window.pwned=1</script>',
  '[click](javascript:alert%281%29)',
  '<a href="javascript:alert(1)" onclick="alert(1)">click</a>',
  '<svg><g onload="alert(1)"></g></svg>',
  '<iframe srcdoc="<script>alert(1)</script>"></iframe>',
  '<form><input name="geminiApiKey"></form>',
  '<math><mtext><img src=x onerror=alert(1)></mtext></math>',
];
for (const attack of attacks) {
  for (const sender of ['Gemini', 'User', 'System', 'Error']) {
    assertSafe(addMessageToChat(sender, attack));
  }
  const updated = addMessageToChat('System', 'Starting');
  updateSystemMessage(updated, attack);
  assertSafe(updated);
}
const sender = '<img src=x onerror=alert(1)>';
const label = addMessageToChat(sender, '**safe**');
assert.equal(label.querySelector('strong').textContent, `${sender}:`);
assertSafe(label);
const benign = addMessageToChat('Gemini', '**bold** and [docs](https://example.com)\n\n- item');
assert.equal(benign.querySelector(':scope > div strong').textContent, 'bold');
assert.equal(benign.querySelector('a').getAttribute('href'), 'https://example.com');
assert.equal(benign.querySelector('li').textContent, 'item');
const long = addMessageToChat('System', attacks[0].repeat(8));
assertSafe(long);
long.querySelector('button').click();
assertSafe(long);

const html = readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf8');
const page = new JSDOM(html).window.document;
const meta = page.querySelector('meta[http-equiv="Content-Security-Policy"]');
assert.ok(meta);
const directives = new Map(meta.content.split(';').map(value => {
  const [name, ...sources] = value.trim().split(/\s+/);
  return [name, sources];
}));
assert.deepEqual(directives.get('script-src'), ["'self'", 'https://appsforoffice.microsoft.com']);
assert.deepEqual(directives.get('object-src'), ["'none'"]);
assert.deepEqual(directives.get('base-uri'), ["'none'"]);
assert.ok(meta.compareDocumentPosition(page.querySelector('script')) & dom.window.Node.DOCUMENT_POSITION_FOLLOWING);
console.log('PASS chat UI sanitization and CSP regressions');
