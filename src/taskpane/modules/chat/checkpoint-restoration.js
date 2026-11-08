import { areRevisionTokensEqual } from '@ansonlai/docx-redline-js';
import { inspectWordSource } from '../docx-redline-js-integration/consumer-core.js';
import { withNativeTrackingDisabled } from '../docx-redline-js-integration/word-ooxml.js';
import { computeWordRestorationToken, computeWordRevertGuardToken } from '../docx-redline-js-integration/word-source-identity.js';
import { queueHostParagraphCount } from './word-document-source.js';

export const CHANGED_SINCE_EDIT_MESSAGE = 'The document changed after this edit. Reverting will also undo those later changes, including any of your own typing.';

/**
 * Restore one body Flat OPC checkpoint once, preserving the host tracking mode.
 *
 * `guardToken` is the Revert guard token (computeWordRevertGuardToken) of the
 * document state the user agreed to discard. When it is passed (a string, or
 * null when unknown), the current body is read first and nothing is written
 * unless it still matches: the result then carries `changedSinceEdit: true` and
 * the current state's `currentToken`, so the caller can ask the user and, on a
 * yes, discard exactly that state and no later one. Leave it undefined to skip
 * the check.
 */
export async function restoreBodyCheckpoint(ooxml, { run = callback => Word.run(callback), getDocumentId = () => null, signal, guardToken } = {}) {
  let written = false, writeAttempted = false;
  try {
    if (signal?.aborted) throw new Error('Checkpoint restoration was cancelled before writing.');
    if (typeof ooxml !== 'string' || !ooxml.trim()) throw new Error('The checkpoint is empty.');
    const documentId = getDocumentId();
    const expected = inspectWordSource(ooxml);
    const before = await computeWordRestorationToken(expected.parts);
    return await run(async context => {
      const body = context.document.body;
      if (guardToken !== undefined) {
        const current = body.getOoxml();
        const readCurrentCount = queueHostParagraphCount(body);
        await context.sync();
        const currentToken = await computeWordRevertGuardToken(inspectWordSource(current.value).parts, { hostParagraphCount: readCurrentCount() });
        if (typeof guardToken !== 'string' || !areRevisionTokensEqual(guardToken, currentToken.value)) {
          return { success: false, verified: false, written: false, writeAttempted: false, changedSinceEdit: true, currentToken: currentToken.value, message: CHANGED_SINCE_EDIT_MESSAGE };
        }
      }
      await withNativeTrackingDisabled(context, async () => {
        if (signal?.aborted) throw new Error('Checkpoint restoration was cancelled before writing.');
        if (getDocumentId() !== documentId) throw new Error('The active document changed before restoration.');
        writeAttempted = true;
        body.insertOoxml(ooxml, 'Replace');
        await context.sync();
        written = true;
      });
      const fresh = body.getOoxml();
      const readCount = queueHostParagraphCount(body);
      await context.sync();
      const hostParagraphCount = readCount();
      const actual = inspectWordSource(fresh.value);
      const after = await computeWordRestorationToken(actual.parts, { hostParagraphCount });
      const verified = getDocumentId() === documentId && areRevisionTokensEqual(before.value, after.value)
        && JSON.stringify(before.commentIdentity) === JSON.stringify(after.commentIdentity);
      return { success: verified, verified, written, writeAttempted,
        message: verified ? 'Reverted and verified the body checkpoint.'
          : 'The checkpoint was inserted, but its restored source could not be verified. Inspect the document before continuing.' };
    });
  } catch (error) {
    return { success: false, verified: false, written, writeAttempted,
      message: written || writeAttempted
        ? 'Checkpoint restoration was attempted but could not be confirmed. Inspect the document before continuing; do not replay it.'
        : `The checkpoint could not be restored. ${error.message}` };
  }
}
