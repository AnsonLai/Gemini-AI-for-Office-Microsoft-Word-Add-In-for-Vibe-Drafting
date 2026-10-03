import { areRevisionTokensEqual } from '@ansonlai/docx-redline-js';
import { computeWordSourceRevisionToken } from '../docx-redline-js-integration/word-source-identity.js';
import { applyOperationsToDocumentXml } from '@ansonlai/docx-redline-js/standalone-runner';
import { executePureOoxmlBatch, inspectWordSource } from '../docx-redline-js-integration/word-operation-runner.js';
import { compileDocumentEdits } from './document-edit-intents.js';
import { queueHostParagraphCount } from './word-document-source.js';

const fail = (code, message) => Object.assign(new Error(message), { code });
function sourceParts(source) {
  const parts = Object.fromEntries(['documentXml', 'commentsXml', 'commentsExtendedXml', 'commentsIdsXml', 'commentsExtensibleXml', 'numberingXml', 'stylesXml'].map(key => [key, source[key] || null]));
  parts.additionalParts = source.additionalParts || Object.fromEntries([
    ['word/commentsIds.xml', parts.commentsIdsXml], ['word/commentsExtensible.xml', parts.commentsExtensibleXml]
  ].filter(([, xml]) => xml));
  return parts;
}
const commentIdentity = comment => [comment.author, comment.text, comment.anchoredText || ''];
function reviewSegments(segments = []) {
  return segments.filter(segment => segment.kind !== 'baseline').reduce((result, segment) => {
    const next = { kind: segment.kind, author: segment.author || null, text: segment.text,
      acceptedStart: segment.acceptedStart, rejectedStart: segment.rejectedStart };
    const previous = result.at(-1);
    const adjacent = key => previous[key] == null && next[key] == null
      || previous[key] != null && previous[key] + previous.text.length === next[key];
    if (previous && previous.kind === next.kind && previous.author === next.author
      && adjacent('acceptedStart') && adjacent('rejectedStart')) previous.text += next.text;
    else result.push(next);
    return result;
  }, []);
}
const acceptedShape = inspection => ({
  paragraphs: inspection.paragraphs.map(paragraph => ({
    text: paragraph.exactText, inTable: paragraph.inTable,
    table: paragraph.table ? [paragraph.table.rowIndex, paragraph.table.cellIndex] : null,
    list: paragraph.list ? [paragraph.list.format, paragraph.list.level, paragraph.list.label] : null,
    review: { present: paragraph.hasRevisions, authors: [...paragraph.revisionAuthors].sort(),
      segments: reviewSegments(paragraph.segments) },
  })),
  comments: (inspection.comments || []).map(comment => {
    const parent = (inspection.comments || []).find(item => comment.parentCommentId != null && String(item.id) === String(comment.parentCommentId));
    return [...commentIdentity(comment), comment.done === true,
      parent ? commentIdentity(parent) : comment.parentCommentId != null || comment.parentParaId ? 'unresolved parent' : null];
  }).sort(),
});

/** Direct document tools. Word is only a source/commit/state transport. */
export function createDocumentTools({ session, getDocumentId, getAuthor, getGenerateRedlines, persistCheckpoint, signal,
  run = callback => Word.run(callback), onReceipt = () => {} }) {
  const checkCancelled = () => { if (signal?.aborted) throw fail('REQUEST_CANCELLED', 'Request cancelled before editing.'); };
  return {
    inspectDocument: options => session.inspect(options),
    async applyDocumentEdits({ edits }) {
      let compiled, expected, result, settledBatch;
      let counts = { sourceReads: 0, contentWrites: 0, verificationReads: 0 };
      try {
        checkCancelled();
        const documentId = getDocumentId();
        const author = getAuthor(), generateRedlines = getGenerateRedlines();
        const ownAuthor = String(author).trim().toLowerCase();
        result = await run(async context => {
          const body = context.document.body;
          let hostParagraphCount;
          let readHostCount;
          const scope = {
            getOoxml() { counts.sourceReads++; readHostCount = queueHostParagraphCount(body); return body.getOoxml(); },
            insertOoxml(...args) { counts.contentWrites++; return body.insertOoxml(...args); },
          };
          const batch = await executePureOoxmlBatch(context, scope, async source => {
            checkCancelled();
            hostParagraphCount = readHostCount();
            if (getDocumentId() !== documentId) throw fail('DOCUMENT_IDENTITY_CHANGED', 'The active document changed. Inspect it again.');
            compiled = await compileDocumentEdits(edits, { session, source: {
              documentId, revisionView: 'accepted', parts: sourceParts(source), hostParagraphCount,
            } });
            compiled.operations = compiled.operations.map(operation =>
              operation.structuredContent === false && ['replace', 'redline'].includes(operation.type)
                && source.paragraphs.some(paragraph => paragraph.index === operation.target.index && paragraph.hasRevisions
                  && (!paragraph.revisionAuthors.length
                    || paragraph.revisionAuthors.some(name => name.trim().toLowerCase() !== ownAuthor)))
                ? { ...operation, existingRevisions: 'slice-cross-author' } : operation);
            return compiled.operations;
          }, {
            author, generateRedlines,
            disableNativeTracking: true, strictTargets: true, expectedRevision: true, signal,
            onPrepared(preparation) {
              if (['ready', 'noop'].includes(preparation.status)
                && (preparation.result?.results?.length !== compiled.operations.length
                  || preparation.result?.receipts?.length !== compiled.operations.length
                  || preparation.result.results.some(item => !['applied', 'no_change'].includes(item.status))
                  || new Set(preparation.result.results.map(item => item.index)).size !== compiled.operations.length)) {
                throw fail('INCOMPLETE_ENGINE_RESULT', 'The engine did not account for every prepared operation. No write was attempted.');
              }
              if (preparation.status === 'ready') {
                expected = inspectWordSource(preparation.insertionPayload);
                const createdIds = new Set(preparation.result.receipts.flatMap(receipt => receipt.revisionItems || [])
                  .map(item => String(item.id)));
                if (generateRedlines === false && expected.inspection.paragraphs.some(paragraph =>
                  paragraph.segments.some(segment => segment.kind !== 'baseline'
                    && String(segment.author || '').trim().toLowerCase() === ownAuthor
                    && createdIds.has(String(segment.revisionId))))) {
                  throw fail('TRACKING_POLICY_MISMATCH', 'The library prepared new tracked text despite tracked edits being disabled. This operation requires a library fix; no part of the batch was written.');
                }
              }
            },
            async beforeWrite(preparation) {
              checkCancelled();
              let checkpointIndex;
              try { checkpointIndex = await persistCheckpoint(preparation.sourceOoxml); }
              catch { throw fail('CHECKPOINT_FAILED', 'The recovery checkpoint could not be saved. No edit was attempted.'); }
              if (!Number.isInteger(checkpointIndex) || checkpointIndex < 1) throw fail('CHECKPOINT_FAILED', 'The recovery checkpoint could not be saved. No edit was attempted.');
              checkCancelled();
              const fresh = body.getOoxml(); counts.sourceReads++;
              const readFreshCount = queueHostParagraphCount(body);
              await context.sync();
              const sourceToken = await computeWordSourceRevisionToken(sourceParts(preparation.source), { hostParagraphCount });
              const freshToken = await computeWordSourceRevisionToken(inspectWordSource(fresh.value).parts, { hostParagraphCount: readFreshCount() });
              if (getDocumentId() !== documentId || !areRevisionTokensEqual(sourceToken.value, freshToken.value)) {
                throw fail('STALE_DOCUMENT_CONTEXT', 'The document changed while saving its checkpoint. Inspect it again before editing.');
              }
              return checkpointIndex;
            },
          });
          settledBatch = batch;
          if (batch.written && batch.status !== 'error') {
            // A settled write is verified even if cancellation arrived during it.
            try {
              const fresh = body.getOoxml(); counts.verificationReads++;
              await context.sync();
              const actual = inspectWordSource(fresh.value);
              if (getDocumentId() !== documentId
                || JSON.stringify(acceptedShape(actual.inspection)) !== JSON.stringify(acceptedShape(expected.inspection))) {
                throw fail('HOST_VERIFICATION_FAILED', 'The committed document does not match the prepared text, structure, review marks or comments.');
              }
              // Reapplying a requested format to the actual source must be a pure no-op.
              // This checks Word's fresh state without issuing another host write.
              for (const expectation of compiled.expectations.filter(item => item.action === 'format_text')) {
                const old = expectation.target;
                const precedingInsertions = compiled.expectations.filter(item => item.target.index < old.index)
                  .reduce((offset, item) => offset + (item.expectedParagraphs ? item.expectedParagraphs.length - 1
                    : item.expectedDeleted && !generateRedlines ? -1 : 0), 0);
                const preparedTarget = expected.inspection.paragraphs.find(paragraph => old.paragraphId && paragraph.paragraphId === old.paragraphId)
                  || expected.inspection.paragraphs.find(paragraph => paragraph.index === old.index + precedingInsertions
                    && paragraph.exactText === expectation.sourceText);
                const target = actual.inspection.paragraphs.find(paragraph => paragraph.index === preparedTarget?.index);
                if (!target) throw fail('HOST_VERIFICATION_FAILED', 'The formatted target could not be verified.');
                const verification = await applyOperationsToDocumentXml(actual.parts.documentXml, [{
                  type: 'format', target, ...expectation.format,
                }], author, actual.parts, { atomic: true, strictTargets: true, generateRedlines: false });
                if (verification.status !== 'ok' || verification.hasChanges) throw fail('HOST_VERIFICATION_FAILED', 'Word did not retain the requested formatting.');
              }
              return { ...batch, verified: true };
            } catch (error) {
              return { ...batch, status: 'error', verified: false, requiresInspection: true, mutationOutcome: 'applied_with_host_error', error: { code: error.code || 'HOST_VERIFICATION_FAILED', message: error.message } };
            }
          }
          return { ...batch, verified: batch.mutationOutcome === 'noop' && batch.status === 'ok' };
        });
      } catch (error) {
        result = { ...settledBatch, status: 'error', verified: false,
          written: settledBatch?.written === true, writeAttempted: settledBatch?.writeAttempted === true || counts.contentWrites > 0,
          mutationOutcome: settledBatch?.written ? 'applied_with_host_error' : settledBatch?.writeAttempted || counts.contentWrites > 0 ? 'indeterminate' : 'refused', error: {
          code: ['STALE_DOCUMENT_SOURCE', 'STALE_TARGET_DESCRIPTOR', 'DOCUMENT_IDENTITY_CHANGED', 'STALE_OR_UNKNOWN_HANDLE'].includes(error.code) ? 'STALE_DOCUMENT_CONTEXT'
            : error.code || (settledBatch?.written ? 'WORD_RUN_POST_COMMIT_ERROR' : counts.contentWrites > 0 ? 'WORD_HOST_OUTCOME_UNCERTAIN' : 'DOCUMENT_BATCH_FAILED'), message: error.message,
          ...(error.recovery ? { recovery: error.recovery } : {}),
        } };
      }
      const success = result.status === 'ok' && result.verified === true;
      if (result.written || result.writeAttempted || result.error?.code?.startsWith('STALE')) session.invalidate();
      const receipt = {
        status: success ? 'ok' : 'error', success, completion: success, verified: result.verified === true,
        written: result.written === true, writeAttempted: result.writeAttempted === true,
        mutationOutcome: result.mutationOutcome, checkpointIndex: result.checkpointIndex,
        requiresInspection: result.requiresInspection === true,
        requestedCount: edits?.length || 0, accountedCount: success ? compiled.requestedCount : 0,
        locations: compiled?.locations?.map(item => item.location) || [], counters: counts,
        ...(result.error ? { error: result.error } : {}),
        message: success ? (result.written ? `Applied and verified ${compiled.requestedCount} requested edit(s).` : 'The requested state is already satisfied; no document write was needed.')
          : result.written || result.writeAttempted ? 'An edit was attempted but completion could not be verified. Inspect the document before continuing; do not replay it.'
            : `No changes made. ${result.error?.message || 'The batch was refused.'}`,
      };
      try { onReceipt(receipt); } catch { /* Diagnostics cannot alter host outcomes. */ }
      return receipt;
    },
  };
}
