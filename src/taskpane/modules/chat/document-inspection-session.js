import {
  areRevisionTokensEqual,
  computeDocumentPartsRevisionToken,
  inspectDocumentParts
} from '@ansonlai/docx-redline-js';

const DEFAULT_MAX_RESULTS = 8;
const MAX_RESULTS_LIMIT = 12;
const DEFAULT_MAX_TEXT_CHARS = 1500;
const MAX_TEXT_CHARS_LIMIT = 4000;
const DEFAULT_MAX_TOTAL_TEXT_CHARS = 12000;
const MAX_TOTAL_TEXT_CHARS_LIMIT = 12000;
const MAX_SEARCH_CHARS = 1000;
const MAX_STORED_HANDLES = 400;
const MAX_STORED_CURSORS = 100;

// Word's WebView2 (served over HTTPS) and Node both provide crypto.randomUUID.
const opaqueId = prefix => `${prefix}_${globalThis.crypto.randomUUID()}`;

function sessionError(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function errorResult(error) {
  return {
    status: 'error',
    error: {
      code: error?.code || 'INSPECTION_SESSION_FAILED',
      message: error?.message || String(error)
    }
  };
}

function assertPositiveInteger(value, label) {
  if (!Number.isInteger(value) || value < 1) {
    throw sessionError('INVALID_INSPECTION_OPTIONS', `${label} must be a positive integer.`);
  }
}

function normalizeRange(range) {
  if (range == null) return null;
  const start = Array.isArray(range) ? range[0] : range.start;
  const end = Array.isArray(range) ? range[1] : range.end;
  assertPositiveInteger(start, 'Range start');
  assertPositiveInteger(end, 'Range end');
  if (end < start) throw sessionError('INVALID_INSPECTION_OPTIONS', 'Range end must be at or after range start.');
  return { start, end };
}

function sameToken(left, right) {
  return left?.scope === right?.scope
    && left?.algorithm === right?.algorithm
    && left?.version === right?.version
    && areRevisionTokensEqual(left?.value, right?.value);
}

function sameSnapshot(left, right) {
  return left.documentId === right.documentId
    && left.revisionView === right.revisionView
    && Object.is(left.revision, right.revision)
    && sameToken(left.revisionToken, right.revisionToken);
}

function snapshotIdentity(snapshot) {
  return {
    documentId: snapshot.documentId,
    revision: snapshot.revision,
    revisionView: snapshot.revisionView,
    revisionToken: {
      algorithm: snapshot.revisionToken.algorithm,
      version: snapshot.revisionToken.version,
      scope: snapshot.revisionToken.scope,
      value: snapshot.revisionToken.value
    }
  };
}

function describeLocation(paragraph) {
  const reference = String(paragraph?.humanReference || paragraph?.excerpt || '')
    .replace(/\[\s*P\d+(?:\|[^\]]*)?\s*\]/gi, ' ')
    .replace(/\bP\d+\b/gi, ' ')
    .replace(/\s+/g, ' ').trim();
  return reference.length > 240 ? `${reference.slice(0, 237)}…` : reference || 'Document paragraph';
}

function takeText(value, budget, cap) {
  const text = String(value || '');
  const result = text.slice(0, Math.max(0, Math.min(cap, budget.remaining)));
  budget.remaining -= result.length;
  return result;
}

function compactComments(comments, paragraphIndex, budget) {
  const anchored = (comments || []).filter(comment => comment?.paragraphIndex === paragraphIndex);
  const matching = anchored.slice(0, 2);
  const compact = matching.map(comment => {
    const text = String(comment.text || '');
    const author = takeText(comment.author || '', budget, 48);
    const shownText = takeText(text, budget, 120);
    return {
      author: author || null,
      text: shownText,
      ...(shownText.length < text.length ? { textTruncated: true, textLength: text.length } : {}),
      resolved: comment.done === true,
      isReply: !!comment.parentCommentId
    };
  });
  if (anchored.length > matching.length) compact.push({ omitted: anchored.length - matching.length });
  return compact;
}

function compactSelection(selection, nextCursor) {
  if (!selection) return null;
  return {
    totalMatches: selection.totalMatches,
    returnedMatches: selection.returnedMatches,
    returnedParagraphs: selection.returnedParagraphs,
    truncated: selection.truncated,
    ...(selection.limit != null ? { limit: selection.limit } : {}),
    ...(selection.around != null ? { around: selection.around } : {}),
    ...(nextCursor ? { nextCursor } : {})
  };
}

/**
 * Creates an app-owned handle store for one Word document identity.
 *
 * `readSource` must return a fresh object from the current Word body source:
 * `{ documentId, parts, revision?, revisionView? }`. `parts` is the public
 * `inspectDocumentParts` input shape and must include the document part plus
 * any comment/numbering/style parts used by the host capture. No Word or DOM
 * objects are read or retained by this module.
 */
export function createDocumentInspectionSession({
  readSource,
  revisionView = 'accepted',
  maxResults = DEFAULT_MAX_RESULTS,
  maxTextChars = DEFAULT_MAX_TEXT_CHARS,
  maxTotalTextChars = DEFAULT_MAX_TOTAL_TEXT_CHARS,
  computeRevisionToken = computeDocumentPartsRevisionToken,
  getExcludedParagraphIndex = () => null,
  onStaleSource = null
} = {}) {
  if (typeof readSource !== 'function') throw new TypeError('createDocumentInspectionSession requires readSource().');
  if (!['accepted', 'rejected'].includes(revisionView)) throw new TypeError('revisionView must be accepted or rejected.');
  if (!Number.isInteger(maxResults) || maxResults < 1 || maxResults > MAX_RESULTS_LIMIT) {
    throw new RangeError(`maxResults must be between 1 and ${MAX_RESULTS_LIMIT}.`);
  }
  if (!Number.isInteger(maxTextChars) || maxTextChars < 1 || maxTextChars > MAX_TEXT_CHARS_LIMIT) {
    throw new RangeError(`maxTextChars must be between 1 and ${MAX_TEXT_CHARS_LIMIT}.`);
  }
  if (!Number.isInteger(maxTotalTextChars) || maxTotalTextChars < 1 || maxTotalTextChars > MAX_TOTAL_TEXT_CHARS_LIMIT) {
    throw new RangeError(`maxTotalTextChars must be between 1 and ${MAX_TOTAL_TEXT_CHARS_LIMIT}.`);
  }

  const sessionId = opaqueId('session');
  const handles = new Map();
  const cursors = new Map();
  let generation = 0;
  let currentSnapshot = null;
  let diagnosticParts = null;
  let diagnosticHostParagraphCount;
  let nextHandleNumber = 0;

  function invalidate() {
    generation += 1;
    currentSnapshot = null;
    diagnosticParts = null;
    diagnosticHostParagraphCount = undefined;
    handles.clear();
    cursors.clear();
  }

  async function loadSnapshot(source = null) {
    const raw = source || await readSource();
    if (!raw || raw.documentId == null || !raw.parts || typeof raw.parts.documentXml !== 'string') {
      throw sessionError('MISSING_CANONICAL_SOURCE', 'A fresh canonical Word source and document identity are required.');
    }
    if (raw.revisionView != null && raw.revisionView !== revisionView) {
      invalidate();
      throw sessionError('REVISION_VIEW_MISMATCH', `The host source is not in the ${revisionView} revision view.`);
    }

    const revisionToken = await computeRevisionToken(raw.parts, raw);
    if (!revisionToken?.coveredParts?.includes('word/document.xml')) {
      throw sessionError('MISSING_CANONICAL_SOURCE', 'The canonical source token does not cover word/document.xml.');
    }

    return {
      raw,
      documentId: raw.documentId,
      revision: raw.revision,
      revisionView,
      parts: raw.parts,
      excludedParagraphIndex: getExcludedParagraphIndex(raw.parts, raw),
      revisionToken
    };
  }

  function activateSnapshot(snapshot) {
    if (currentSnapshot && !sameSnapshot(currentSnapshot, snapshot)) invalidate();
    if (!currentSnapshot) currentSnapshot = snapshotIdentity(snapshot);
    if (onStaleSource) { diagnosticParts = snapshot.parts; diagnosticHostParagraphCount = snapshot.raw.hostParagraphCount; }
    return currentSnapshot;
  }

  function makeHandle(kind, paragraph, snapshot, extras = {}) {
    nextHandleNumber += 1;
    const handle = `${sessionId}_h${nextHandleNumber.toString(36)}`;
    handles.set(handle, {
      kind,
      index: paragraph.index,
      documentId: snapshot.documentId,
      revision: snapshot.revision,
      revisionView: snapshot.revisionView,
      revisionToken: snapshot.revisionToken,
      generation,
      ...extras
    });
    return handle;
  }

  function libraryOptions({ search, range, around, limit, after }) {
    return {
      revisionView,
      excerptLength: Math.min(maxTextChars, 240),
      ...(search ? { search } : {}),
      ...(range ? { range } : {}),
      ...(around ? { around } : {}),
      ...(after != null ? { after } : {}),
      limit
    };
  }

  function normalizeInspectOptions(options = {}) {
    const cursor = options.cursor == null ? null : String(options.cursor);
    if (cursor) {
      if (options.search != null || options.range != null || options.around != null || options.limit != null || options.includeDocumentEnd) {
        throw sessionError('INVALID_INSPECTION_OPTIONS', 'A cursor continues its original query; do not combine it with new inspection options.');
      }
      const saved = cursors.get(cursor);
      if (!saved || saved.generation !== generation) throw sessionError('STALE_CURSOR', 'This inspection cursor has expired. Inspect the current source again.');
      const continued = { ...saved.options, after: saved.after };
      return { ...continued, cursor, libraryOptions: libraryOptions(continued) };
    }

    const search = options.search == null ? null : String(options.search).trim();
    if (search != null && (!search || search.length > MAX_SEARCH_CHARS)) {
      throw sessionError('INVALID_INSPECTION_OPTIONS', `Search must contain 1 to ${MAX_SEARCH_CHARS} characters.`);
    }
    const range = normalizeRange(options.range);
    const around = options.around == null ? 0 : options.around;
    if (!Number.isInteger(around) || around < 0 || around > 5) {
      throw sessionError('INVALID_INSPECTION_OPTIONS', 'around must be an integer from 0 through 5.');
    }
    const limit = options.limit == null ? maxResults : options.limit;
    if (!Number.isInteger(limit) || limit < 1 || limit > maxResults) {
      throw sessionError('INVALID_INSPECTION_OPTIONS', `limit must be between 1 and ${maxResults}.`);
    }
    const query = { search, range, around, limit };
    return { ...query, includeDocumentEnd: options.includeDocumentEnd === true, libraryOptions: libraryOptions(query) };
  }

  function targetView(paragraph, handle, inspection, handleByIndex, totalBudget, remainingTargets) {
    const exactText = String(paragraph.exactText || '');
    const slotBudget = { remaining: Math.floor(totalBudget.remaining / Math.max(1, remainingTargets)) };
    const metadataReserve = Math.min(240, Math.floor(slotBudget.remaining * 0.16));
    const text = takeText(exactText, slotBudget, Math.min(maxTextChars, slotBudget.remaining - metadataReserve));
    const location = takeText(describeLocation(paragraph), slotBudget, Math.min(160, metadataReserve));
    const comments = compactComments(inspection.comments, paragraph.index, slotBudget);
    const cues = {
      inTable: paragraph.inTable === true,
      ...(paragraph.headingLevel ? { headingLevel: paragraph.headingLevel } : {}),
      ...(paragraph.list ? {
        list: {
          ...(paragraph.list.label ? { label: takeText(paragraph.list.label, slotBudget, 40) } : {}),
          ...(paragraph.list.format ? { format: takeText(paragraph.list.format, slotBudget, 40) } : {}),
          level: paragraph.list.level
        }
      } : {}),
      hasRevisions: paragraph.hasRevisions === true,
      ...(paragraph.revisionAuthors?.length ? {
        revisionAuthors: paragraph.revisionAuthors.slice(0, 2).map(author => takeText(author, slotBudget, 48)).filter(Boolean)
      } : {}),
      ...(comments.length ? { comments } : {})
    };
    const used = Math.floor(totalBudget.remaining / Math.max(1, remainingTargets)) - slotBudget.remaining;
    totalBudget.remaining -= used;
    return {
      textComplete: text.length === exactText.length,
      target: {
        handle,
        location,
        text,
        ...(text.length < exactText.length ? { textTruncated: true, textLength: exactText.length } : {}),
        role: paragraph.selectionRole || 'match',
        ...(Array.isArray(paragraph.contextFor) && paragraph.contextFor.length ? {
          contextFor: paragraph.contextFor.map(index => handleByIndex.get(index)).filter(Boolean)
        } : {}),
        cues
      }
    };
  }

  async function inspect(options = {}) {
    try {
      const normalized = normalizeInspectOptions(options);
      const snapshot = await loadSnapshot();
      if (normalized.cursor) {
        const cursorSnapshot = cursors.get(normalized.cursor);
        if (!cursorSnapshot || cursorSnapshot.generation !== generation
          || !sameSnapshot(cursorSnapshot.snapshot, snapshot)) {
          invalidate();
          throw sessionError('STALE_CURSOR', 'The document changed while paging this inspection. Inspect the current source again.');
        }
      }
      activateSnapshot(snapshot);

      const selectedInspection = inspectDocumentParts(snapshot.parts, normalized.libraryOptions);
      if (selectedInspection?.status !== 'ok') {
        throw sessionError(selectedInspection?.error?.code || 'CANONICAL_INSPECTION_FAILED', selectedInspection?.error?.message || 'The Word source could not be inspected.');
      }

      // Library inspection always returns strong descriptors (index, exactText, paragraphId, fingerprint).
      const editableParagraphs = (selectedInspection.paragraphs || []).filter(paragraph => paragraph.index !== snapshot.excludedParagraphIndex);
      if (handles.size + editableParagraphs.length + (normalized.includeDocumentEnd ? 1 : 0) > MAX_STORED_HANDLES) {
        invalidate();
        throw sessionError('HANDLE_STORE_LIMIT', 'This inspection session reached its handle limit. Start a fresh inspection session.');
      }

      const issued = [];
      const handleByIndex = new Map();
      for (const paragraph of editableParagraphs) {
        const handle = makeHandle('paragraph', paragraph, snapshot);
        issued.push({ paragraph, handle });
        handleByIndex.set(paragraph.index, handle);
      }

      let nextCursor = null;
      const nextAfter = selectedInspection.selection?.nextAfter;
      if (selectedInspection.selection?.truncated && Number.isInteger(nextAfter) && issued.length) {
        if (cursors.size >= MAX_STORED_CURSORS) {
          invalidate();
          throw sessionError('CURSOR_STORE_LIMIT', 'This inspection session reached its pagination limit. Inspect a narrower range.');
        }
        nextCursor = opaqueId('cursor');
        const savedOptions = {
          search: normalized.search,
          range: normalized.range,
          around: normalized.around,
          limit: normalized.limit
        };
        cursors.set(nextCursor, {
          generation,
          snapshot: snapshotIdentity(snapshot),
          options: savedOptions,
          after: nextAfter
        });
      }

      const totalBudget = { remaining: maxTotalTextChars };
      const projected = issued.map(({ paragraph, handle }, index) => targetView(
        paragraph,
        handle,
        selectedInspection,
        handleByIndex,
        totalBudget,
        issued.length - index
      ));
      for (let index = 0; index < issued.length; index++) {
        handles.get(issued[index].handle).fullTextInspected = projected[index].textComplete;
      }


      let endHandle;
      if (normalized.includeDocumentEnd) {
        const fullInspection = inspectDocumentParts(snapshot.parts, {
          revisionView,
          excerptLength: Math.min(maxTextChars, 240)
        });
        const lastParagraph = fullInspection.paragraphs?.filter(paragraph => paragraph.index !== snapshot.excludedParagraphIndex).at(-1);
        if (fullInspection.status === 'ok' && lastParagraph) {
          endHandle = makeHandle('document-end', lastParagraph, snapshot, {
            fullTextInspected: false
          });
        }
      }

      const targets = projected.map(item => item.target);
      return {
        status: 'ok',
        revisionView,
        targets,
        ...(selectedInspection.selection ? { selection: compactSelection(selectedInspection.selection, nextCursor) } : {}),
        ...(endHandle ? { documentEnd: { handle: endHandle, location: 'End of document' } } : {})
      };
    } catch (error) {
      return errorResult(error);
    }
  }

  async function resolveHandles(handleList, options = {}) {
    try {
      if (!Array.isArray(handleList) || handleList.length === 0) {
        throw sessionError('INVALID_HANDLES', 'At least one inspected handle is required.');
      }
      const entries = handleList.map(handle => {
        const entry = handles.get(handle);
        if (!entry || entry.generation !== generation) throw sessionError('STALE_OR_UNKNOWN_HANDLE', 'This document handle is unknown or expired. Inspect the current source again.');
        return { handle, entry };
      });
      const expected = entries[0].entry;
      if (entries.some(({ entry }) => entry.documentId !== expected.documentId
        || entry.revision !== expected.revision
        || entry.revisionView !== expected.revisionView
        || !sameToken(entry.revisionToken, expected.revisionToken))) {
        throw sessionError('CROSS_SNAPSHOT_HANDLE', 'Handles from different document snapshots cannot be resolved together.');
      }

      const snapshot = await loadSnapshot(options.source || null);
      if (snapshot.documentId !== expected.documentId) {
        invalidate();
        throw sessionError('DOCUMENT_IDENTITY_CHANGED', 'The active Word document changed. Inspect it again before editing.');
      }
      if (snapshot.revisionView !== expected.revisionView || !sameToken(snapshot.revisionToken, expected.revisionToken)
        || !Object.is(snapshot.revision, expected.revision)) {
        try { onStaleSource?.({ before: diagnosticParts, after: snapshot.parts,
          beforeHostParagraphCount: diagnosticHostParagraphCount, afterHostParagraphCount: snapshot.raw.hostParagraphCount,
          excludedParagraphIndex: snapshot.excludedParagraphIndex,
        }); } catch { /* Diagnostics only. */ }
        invalidate();
        throw sessionError('STALE_DOCUMENT_SOURCE', 'The Word source changed after inspection. Inspect the current document again before editing.');
      }

      // An equal revision token means the canonical source is unchanged, so each handle's
      // descriptor still resolves. The apply path resolves it again strictly in the library
      // (strictTargets + expectedRevision); no local descriptor comparison is needed.
      const indexes = [...new Set(entries.map(({ entry }) => entry.index))];
      const freshInspection = inspectDocumentParts(snapshot.parts, { revisionView, indexes });
      if (freshInspection?.status !== 'ok') {
        throw sessionError(freshInspection?.error?.code || 'CANONICAL_INSPECTION_FAILED', freshInspection?.error?.message || 'The current Word source could not be inspected.');
      }
      const freshByIndex = new Map((freshInspection.paragraphs || []).map(item => [item.index, item]));
      const targets = entries.map(({ handle, entry }) => {
        const current = freshByIndex.get(entry.index);
        if (!current) {
          invalidate();
          throw sessionError('STALE_TARGET_DESCRIPTOR', 'The inspected target no longer matches the current canonical source.');
        }
        return {
          handle,
          kind: entry.kind,
          ...(entry.kind === 'document-end' ? { position: 'after' } : {}),
          fullTextInspected: entry.fullTextInspected === true,
          location: describeLocation(current),
          descriptor: current
        };
      });

      return {
        status: 'ok',
        source: {
          ...snapshot.raw,
          documentId: snapshot.documentId,
          revision: snapshot.revision,
          revisionView,
          revisionToken: snapshot.revisionToken
        },
        targets
      };
    } catch (error) {
      return errorResult(error);
    }
  }

  return Object.freeze({
    inspect,
    resolveHandles,
    invalidate
  });
}
