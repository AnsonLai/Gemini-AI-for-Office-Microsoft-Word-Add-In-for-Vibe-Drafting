import assert from 'node:assert/strict';
import { JSDOM } from 'jsdom';

const dom = new JSDOM('<div id="chat-messages"></div>');
globalThis.window = dom.window;
globalThis.document = dom.window.document;
const {
  addConfirmButtons, addMessageToChat, createTypingIndicator, registerChatUiHandlers, removeCancelButton, updateSystemMessage,
} = await import('../src/taskpane/modules/chat/chat-ui.js');

// Cancel stays available while a request's status text changes, and goes when the request ends.
{
  let cancels = 0;
  registerChatUiHandlers({ onCancelRequest: () => { cancels++; } });
  const bubble = createTypingIndicator('teal', true);
  document.getElementById('chat-messages').appendChild(bubble);
  updateSystemMessage(bubble, 'Inspecting the document...');
  assert.equal(bubble.querySelector('.typing-indicator'), null, 'status text replaces the dots');
  const cancel = bubble.querySelector('.cancel-request-btn');
  assert.ok(cancel, 'Cancel survives the first status update');
  updateSystemMessage(bubble, 'Applied and verified 1 requested edit(s).', 7);
  assert.ok(bubble.querySelector('.cancel-request-btn'), 'Cancel survives a receipt with a Revert button');
  assert.ok(bubble.querySelector('.revert-checkpoint-btn'), 'the receipt still gets its Revert button');
  updateSystemMessage(bubble, 'Searching the web...');
  assert.ok(bubble.querySelector('.revert-checkpoint-btn'), 'a later status keeps the Revert button');
  bubble.querySelector('.cancel-request-btn').click();
  assert.equal(cancels, 1, 'the kept button still cancels');
  removeCancelButton(bubble);
  assert.equal(bubble.querySelector('.cancel-request-btn'), null, 'a finished request has no Cancel button');
  assert.ok(bubble.querySelector('.revert-checkpoint-btn'), 'removing Cancel keeps the Revert button');
  removeCancelButton(null);
}

// A plain status message never gains a Cancel button.
{
  const message = addMessageToChat('System', 'Reverting to checkpoint...');
  updateSystemMessage(message, 'Reverted and verified the body checkpoint.');
  assert.equal(message.querySelector('.cancel-request-btn'), null);
}

// A confirmation stays open when its handler refuses, and closes once it runs.
{
  const message = addMessageToChat('System', 'The document changed after this edit.');
  let confirmations = 0, allow = false;
  addConfirmButtons(message, {
    confirmLabel: 'Revert anyway', cancelLabel: 'Keep document',
    onConfirm: () => { confirmations++; return allow ? undefined : false; },
  });
  const confirm = () => [...message.querySelectorAll('button')].find(button => button.textContent === 'Revert anyway');
  confirm().click();
  assert.equal(confirmations, 1);
  assert.ok(confirm(), 'a refused confirmation can be pressed again later');
  allow = true;
  confirm().click();
  assert.equal(confirmations, 2);
  assert.equal(confirm(), undefined, 'an accepted confirmation closes');
}

{
  const message = addMessageToChat('System', 'The document changed after this edit.');
  let kept = false;
  addConfirmButtons(message, { confirmLabel: 'Revert anyway', cancelLabel: 'Keep document', onConfirm: () => {}, onCancel: () => { kept = true; } });
  [...message.querySelectorAll('button')].find(button => button.textContent === 'Keep document').click();
  assert.equal(kept, true);
  assert.equal(message.querySelector('button'), null, '"Keep document" closes the question');
}

console.log('PASS: chat UI keeps Cancel through status updates and keeps refused confirmations open');
