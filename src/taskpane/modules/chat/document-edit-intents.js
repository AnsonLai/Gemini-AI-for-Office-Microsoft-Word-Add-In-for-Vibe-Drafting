import { analyzeStructuredContent, preprocessMarkdown } from '@ansonlai/docx-redline-js';

const TEXT_ACTIONS = new Set(['replace_text', 'rewrite_paragraph', 'delete_paragraph', 'insert_paragraph', 'append_paragraph']);
const FORMAT_KEYS = new Set(['bold', 'italic', 'underline', 'strikethrough', 'highlight']);
const FIELDS = {
  replace_text: ['target', 'find', 'replace', 'occurrence'],
  rewrite_paragraph: ['target', 'text'],
  delete_paragraph: ['target'],
  insert_paragraph: ['target', 'text', 'position'],
  append_paragraph: ['target', 'text'],
  comment: ['target', 'text', 'anchor'],
  format_text: ['target', 'find', 'occurrence', 'properties']
};

export class DocumentEditIntentError extends Error {
  constructor(code, message, details = {}) {
    super(message);
    this.name = 'DocumentEditIntentError';
    this.code = code;
    if (details.field) this.field = details.field;
    if (Number.isInteger(details.actionIndex)) this.actionIndex = details.actionIndex;
    this.recovery = {
      action: details.requiresReinspection ? 'reinspect' : 'change_request',
      sameArgumentsSafe: false,
      requiresReinspection: details.requiresReinspection === true,
      requiresUserAuthorization: false
    };
  }
}

function refuse(code, message, details) { throw new DocumentEditIntentError(code, message, details); }
function record(value) { return !!value && typeof value === 'object' && !Array.isArray(value); }
function fail(code, message, index, field, requiresReinspection = false) {
  refuse(code, message, { actionIndex: index, field: index == null ? field : `edits[${index}].${field}`, requiresReinspection });
}
function text(value, index, field, allowEmpty = false) {
  if (typeof value !== 'string' || (!allowEmpty && !value.trim())) fail('INVALID_DOCUMENT_EDIT', `${field} must be ${allowEmpty ? 'a' : 'a non-empty'} string.`, index, field);
  return value;
}

function plainParagraph(value, index, field) {
  if (value.includes('\n') || value.includes('\r') || value !== value.trim()) {
    fail('UNSUPPORTED_DOCUMENT_EDIT', 'Inserted content must be one plain paragraph without line breaks or outer whitespace.', index, field);
  }
  const analysis = analyzeStructuredContent(value);
  const markdown = preprocessMarkdown(value);
  if (!analysis.valid || analysis.blocks.length !== 1 || analysis.blocks[0]?.type !== 'paragraph'
    || markdown.formatHints.length || markdown.cleanText !== value) {
    fail('UNSUPPORTED_DOCUMENT_EDIT', 'Inserted content must be plain text; lists, tables, headings, and inline formatting are unsupported.', index, field);
  }
}

function validateEdit(edit, index) {
  if (!record(edit) || typeof edit.type !== 'string' || !FIELDS[edit.type]) {
    fail('UNSUPPORTED_DOCUMENT_EDIT', 'This edit action is unsupported. Use a supported text, paragraph, comment, or basic-formatting action.', index, 'type');
  }
  const allowed = new Set(['type', ...FIELDS[edit.type]]);
  const extra = Object.keys(edit).find(key => !allowed.has(key));
  if (extra) fail('INVALID_DOCUMENT_EDIT', `Field "${extra}" is not supported for ${edit.type}.`, index, extra);
  const value = { ...edit, target: text(edit.target, index, 'target') };

  if (edit.type === 'replace_text') {
    value.find = text(edit.find, index, 'find');
    value.replace = text(edit.replace, index, 'replace', true);
    if (/[\r\n]/.test(value.find + value.replace)) fail('UNSUPPORTED_DOCUMENT_EDIT', 'Literal replacement cannot span or insert paragraph breaks.', index, 'find');
    if (value.find === value.replace) fail('UNSUPPORTED_LITERAL_NOOP', 'A literal replacement must change the matched text. Use rewrite_paragraph with the complete desired paragraph to confirm an already-satisfied state.', index, 'replace');
  } else if (edit.type === 'rewrite_paragraph') {
    value.text = text(edit.text, index, 'text');
    if (/[\r\n]/.test(value.text)) fail('UNSUPPORTED_DOCUMENT_EDIT', 'A paragraph rewrite must contain one paragraph. Use a paragraph insertion for a new paragraph.', index, 'text');
  } else if (edit.type === 'insert_paragraph' || edit.type === 'append_paragraph') {
    value.text = text(edit.text, index, 'text');
    plainParagraph(value.text, index, 'text');
    if (edit.type === 'insert_paragraph' && !['before', 'after'].includes(edit.position)) {
      fail('INVALID_DOCUMENT_EDIT', 'position must be "before" or "after".', index, 'position');
    }
  } else if (edit.type === 'comment') {
    value.text = text(edit.text, index, 'text');
    if (edit.anchor !== undefined) value.anchor = text(edit.anchor, index, 'anchor');
  } else if (edit.type === 'format_text') {
    value.find = text(edit.find, index, 'find');
    if (/[\r\n]/.test(value.find)) fail('UNSUPPORTED_DOCUMENT_EDIT', 'Formatting must target text within one paragraph.', index, 'find');
    if (edit.occurrence !== undefined && (!Number.isInteger(edit.occurrence) || edit.occurrence < 1)) {
      fail('INVALID_DOCUMENT_EDIT', 'occurrence must be a positive integer.', index, 'occurrence');
    }
    if (!record(edit.properties) || !Object.keys(edit.properties).length) fail('INVALID_DOCUMENT_EDIT', 'properties must contain a supported formatting change.', index, 'properties');
    const bad = Object.keys(edit.properties).find(key => !FORMAT_KEYS.has(key));
    if (bad) fail('UNSUPPORTED_DOCUMENT_EDIT', `Formatting property "${bad}" is outside the supported allowlist.`, index, `properties.${bad}`);
    for (const [key, property] of Object.entries(edit.properties)) {
      if (key === 'highlight' ? !['yellow', 'none'].includes(property) : typeof property !== 'boolean') {
        fail('UNSUPPORTED_DOCUMENT_EDIT', key === 'highlight' ? 'Only yellow highlighting and removing highlighting are supported.' : `${key} must be true or false.`, index, `properties.${key}`);
      }
    }
    value.properties = { ...edit.properties };
  }
  if (edit.occurrence !== undefined && edit.type === 'replace_text'
    && (!Number.isInteger(edit.occurrence) || edit.occurrence < 1)) fail('INVALID_DOCUMENT_EDIT', 'occurrence must be a positive integer.', index, 'occurrence');
  return { edit: value, actionIndex: index };
}

function identity(descriptor) { return descriptor.paragraphId ? `id:${descriptor.paragraphId}` : `${descriptor.index}|${descriptor.fingerprint}`; }
function targetDescriptor(descriptor) {
  return {
    exactText: descriptor.exactText, index: descriptor.index,
    ...(descriptor.paragraphId ? { paragraphId: descriptor.paragraphId } : {}),
    fingerprint: descriptor.fingerprint, revisionView: descriptor.revisionView, inTable: descriptor.inTable
  };
}
function location(target) {
  const descriptor = target.descriptor;
  return String(target.location || descriptor.humanReference || descriptor.excerpt || descriptor.provision || descriptor.nearestHeading?.text || 'Document paragraph')
    .replace(/\s+/g, ' ').trim().slice(0, 240);
}
function checkedDescriptor(target, index, expectedKind) {
  if (target?.kind !== expectedKind) fail('INVALID_DOCUMENT_EDIT_TARGET', `This action requires an inspected ${expectedKind === 'document-end' ? 'document-end' : 'paragraph'} handle.`, index, 'target', true);
  const descriptor = target.descriptor;
  if (!record(descriptor) || !Number.isInteger(descriptor.index) || descriptor.index < 1
    || typeof descriptor.exactText !== 'string' || typeof descriptor.fingerprint !== 'string'
    || !descriptor.fingerprint || typeof descriptor.inTable !== 'boolean') {
    fail('WEAK_DOCUMENT_EDIT_TARGET', 'The target has no strong canonical descriptor. Inspect the current source again.', index, 'target', true);
  }
  return descriptor;
}
function checkedInsertionTarget(descriptor, index) {
  if (descriptor.inTable || descriptor.table || descriptor.list || descriptor.headingLevel
    || (descriptor.styleId && /^heading\b/i.test(String(descriptor.styleId)))) {
    fail('UNSUPPORTED_DOCUMENT_EDIT_TARGET', 'Plain paragraph insertion beside a list, table, or heading is unsupported.', index, 'target');
  }
  const processed = preprocessMarkdown(descriptor.exactText);
  const analysis = analyzeStructuredContent(descriptor.exactText);
  if (descriptor.exactText !== descriptor.exactText.trim() || /[\r\n]/.test(descriptor.exactText)
    || analysis.blocks.length > 1 || analysis.blocks.some(block => block.type !== 'paragraph')
    || processed.formatHints.length || processed.cleanText !== descriptor.exactText) {
    fail('UNSUPPORTED_DOCUMENT_EDIT_TARGET', 'This target has spacing, soft breaks, or structural/formatting markers that cannot be safely preserved by plain insertion.', index, 'target');
  }
}
function resolvedTargets(result, handles) {
  if (result?.status !== 'ok' || !Array.isArray(result.targets) || result.targets.length !== handles.length) {
    refuse(result?.error?.code || 'DOCUMENT_TARGET_RESOLUTION_FAILED', result?.error?.message || 'Targets could not be resolved against the current source. Inspect again.', { requiresReinspection: true });
  }
  const map = new Map();
  for (let i = 0; i < handles.length; i++) {
    const target = result.targets[i];
    if (target?.handle !== handles[i]) refuse('DOCUMENT_TARGET_RESOLUTION_FAILED', 'Resolved targets did not match the requested handles. Inspect again.', { requiresReinspection: true });
    map.set(handles[i], target);
  }
  return map;
}

function combinations(items, targets) {
  const groups = new Map();
  for (const item of items) {
    const target = targets.get(item.edit.target);
    const key = identity(target.descriptor);
    const group = groups.get(key) || { text: [], comments: [], formats: [] };
    if (TEXT_ACTIONS.has(item.edit.type)) group.text.push(item);
    else if (item.edit.type === 'comment') group.comments.push(item);
    else group.formats.push(item);
    groups.set(key, group);
  }
  for (const group of groups.values()) {
    const textActions = group.text.filter(({ edit }) => {
      if (edit.type === 'replace_text') return true;
      const descriptor = targets.get(edit.target).descriptor;
      return !(edit.type === 'rewrite_paragraph' && edit.text === descriptor.exactText);
    });
    const index = item => item.actionIndex;
    if (textActions.length > 1) fail('CONFLICTING_DOCUMENT_EDITS', 'Only one text-changing action can target a source paragraph per batch.', index(textActions[1]), 'target');
    if (group.comments.length > 1) fail('UNSUPPORTED_DOCUMENT_EDIT_COMBINATION', 'Use one comment per source paragraph in a batch.', index(group.comments[1]), 'target');
    if (group.formats.length > 1) fail('UNSUPPORTED_DOCUMENT_EDIT_COMBINATION', 'Combine formatting properties into one format_text action per source paragraph.', index(group.formats[1]), 'target');
    if (textActions.length && group.comments.length) {
      const rewrite = textActions.length === 1 && textActions[0].edit.type === 'rewrite_paragraph'
        ? textActions[0]
        : null;
      if (!rewrite) fail('COMMENT_ANCHOR_REQUIRES_REINSPECTION', 'A text replacement, insertion, or deletion cannot share a source paragraph with a comment. Apply the text change, inspect the result, then add the comment.', index(group.comments[0]), 'anchor', true);
      const paragraph = targets.get(rewrite.edit.target).descriptor.exactText;
      const proposed = rewrite.edit.text;
      const anchor = group.comments[0].edit.anchor || paragraph;
      const uniqueIn = value => value.indexOf(anchor) >= 0 && value.indexOf(anchor) === value.lastIndexOf(anchor);
      if (!uniqueIn(paragraph) || !uniqueIn(proposed)) {
        fail('COMMENT_ANCHOR_REQUIRES_REINSPECTION', 'A rewrite and comment can share a batch only when the original comment anchor remains unique in the proposed paragraph. Otherwise apply the rewrite, inspect the result, then add the comment.', index(group.comments[0]), 'anchor', true);
      }
    }
    if (textActions.length && group.formats.length) fail('UNSUPPORTED_DOCUMENT_EDIT_COMBINATION', 'Formatting and text changes to the same source paragraph cannot share a library batch.', index(group.formats[0]), 'target');
  }
}

/** Resolve opaque handles from one fresh source and compile one all-or-nothing library batch. */
export async function compileDocumentEdits(edits, { session, source } = {}) {
  if (!Array.isArray(edits) || !edits.length) refuse('INVALID_DOCUMENT_EDIT_BATCH', 'edits must be a non-empty array.');
  if (!session || typeof session.resolveHandles !== 'function') refuse('INVALID_DOCUMENT_EDIT_SESSION', 'A document inspection session is required.');
  const items = edits.map(validateEdit);
  const handles = [...new Set(items.map(({ edit }) => edit.target))];
  let result;
  try { result = await session.resolveHandles(handles, { source }); }
  catch (error) { refuse(error?.code || 'DOCUMENT_TARGET_RESOLUTION_FAILED', error?.message || 'Targets could not be resolved. Inspect the current source again.', { requiresReinspection: true }); }
  const targets = resolvedTargets(result, handles);

  for (const { edit, actionIndex } of items) {
    const append = edit.type === 'append_paragraph';
    const target = targets.get(edit.target);
    const descriptor = checkedDescriptor(target, actionIndex, append ? 'document-end' : 'paragraph');
    if (!append && target.fullTextInspected !== true) fail('TARGET_TEXT_NOT_FULLY_INSPECTED', 'This target was only partially shown. Inspect the complete paragraph before editing it.', actionIndex, 'target', true);
    if (append && target.position !== 'after') fail('INVALID_DOCUMENT_EDIT_TARGET', 'The document-end handle does not describe an append position.', actionIndex, 'target', true);
    if (edit.type === 'append_paragraph' || edit.type === 'insert_paragraph') checkedInsertionTarget(descriptor, actionIndex);
    if (edit.type === 'insert_paragraph' && !descriptor.exactText) fail('UNSUPPORTED_DOCUMENT_EDIT_TARGET', 'Insert before or after a blank paragraph is ambiguous; choose a non-empty paragraph.', actionIndex, 'target');
    if (edit.type === 'comment') {
      const anchor = edit.anchor ?? descriptor.exactText;
      if (!anchor || !descriptor.exactText.includes(anchor)) fail('COMMENT_ANCHOR_REQUIRES_REINSPECTION', 'The comment anchor is not present in inspected text. Inspect the resulting paragraph before commenting.', actionIndex, 'anchor', true);
    }
  }
  combinations(items, targets);

  const operations = [], expectations = [], locations = [];
  for (const { edit, actionIndex } of items) {
    const resolved = targets.get(edit.target), descriptor = resolved.descriptor, target = targetDescriptor(descriptor);
    let operation, expected = { action: edit.type, handle: edit.target, target, sourceText: descriptor.exactText };
    switch (edit.type) {
      case 'replace_text': {
        const replacement = { find: edit.find, replace: edit.replace, ...(edit.occurrence ? { occurrence: edit.occurrence } : {}) };
        operation = { type: 'replace', target, replacements: [replacement], structuredContent: false };
        expected.replacement = replacement;
        break;
      }
      case 'rewrite_paragraph':
        if (edit.text === descriptor.exactText) expected.alreadySatisfied = true;
        else operation = { type: 'redline', target, modified: edit.text, structuredContent: false };
        expected.expectedParagraphs = [edit.text];
        break;
      case 'delete_paragraph':
        operation = { type: 'delete', target };
        expected.expectedDeleted = true;
        break;
      case 'insert_paragraph': {
        const current = descriptor.exactText;
        operation = { type: 'redline', target, modified: edit.position === 'before' ? `${edit.text}\n${current}` : `${current}\n${edit.text}`, structuredContent: true };
        expected.expectedParagraphs = edit.position === 'before' ? [edit.text, current] : [current, edit.text];
        break;
      }
      case 'append_paragraph': {
        const current = descriptor.exactText;
        operation = { type: 'redline', target, modified: current ? `${current}\n${edit.text}` : edit.text, structuredContent: true };
        expected.expectedParagraphs = current ? [current, edit.text] : [edit.text];
        break;
      }
      case 'comment': {
        const textToComment = edit.anchor ?? descriptor.exactText;
        operation = { type: 'comment', target, textToComment, commentContent: edit.text };
        expected.comment = { textToComment, commentContent: edit.text };
        break;
      }
      case 'format_text':
        operation = { type: 'format', target, textToFormat: edit.find, ...(edit.occurrence ? { textOccurrence: edit.occurrence } : {}), properties: { ...edit.properties } };
        expected.format = { textToFormat: edit.find, ...(edit.occurrence ? { textOccurrence: edit.occurrence } : {}), properties: { ...edit.properties } };
        break;
      default: fail('UNSUPPORTED_DOCUMENT_EDIT', 'This edit action is unsupported.', actionIndex, 'type');
    }
    if (operation) operations.push(operation);
    expectations.push(expected);
    locations.push({ handle: edit.target, location: location(resolved) });
  }
  return { operations, expectations, locations, requestedCount: edits.length };
}
