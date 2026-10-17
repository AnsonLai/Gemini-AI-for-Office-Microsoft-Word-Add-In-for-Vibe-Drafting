import { areRevisionTokensEqual } from '@ansonlai/docx-redline-js';
import { computeWordRestorationToken, inspectWordFlatOpc } from '@ansonlai/docx-redline-js/word';
import { withNativeTrackingDisabled } from '../docx-redline-js-integration/word-ooxml.js';
import { queueHostParagraphCount } from './word-document-source.js';

/** Restore one body Flat OPC checkpoint once, preserving the host tracking mode. */
export async function restoreBodyCheckpoint(ooxml, { run = callback => Word.run(callback), getDocumentId = () => null, signal } = {}) {
  let written = false, writeAttempted = false;
  try {
    if (signal?.aborted) throw new Error('Checkpoint restoration was cancelled before writing.');
    if (typeof ooxml !== 'string' || !ooxml.trim()) throw new Error('The checkpoint is empty.');
    const documentId = getDocumentId();
    const expected = inspectWordFlatOpc(ooxml);
    const before = await computeWordRestorationToken(expected.parts);
    return await run(async context => {
      const body = context.document.body;
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
      const actual = inspectWordFlatOpc(fresh.value);
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
