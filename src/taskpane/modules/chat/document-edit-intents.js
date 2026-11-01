import { inspectDocumentParts } from '@ansonlai/docx-redline-js';

const TEXT_ACTIONS = new Set(['replace_text', 'rewrite_paragraph', 'delete_paragraph', 'insert_paragraph', 'append_paragraph', 'insert_list_item']);
const FORMAT_KEYS = new Set(['bold', 'italic', 'underline', 'strikethrough', 'highlight']);
const FIELDS = {
  replace_text: ['target', 'find', 'replace', 'occurrence'],
  rewrite_paragraph: ['target', 'text'],
  delete_paragraph: ['target'],
  insert_paragraph: ['target', 'text', 'position'],
  append_paragraph: ['target', 'text'],
  insert_list_item: ['target', 'text'],
  comment: ['target', 'text', 'anchor'],
  format_text: ['target', 'find', 'occurrence', 'properties'],
  convert_to_list: ['target', 'through', 'format', 'markers', 'start'],
  change_list_format: ['target', 'format', 'level', 'start'],
  continue_list: ['target', 'from']
};
const LIST_ACTIONS = new Set(['convert_to_list', 'change_list_format', 'continue_list']);
const LIST_FORMATS = ['decimal', 'upperLetter', 'lowerLetter', 'upperRoman', 'lowerRoman', 'bullet'];

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

function listFormat(value, index, field = 'format') {
  if (!LIST_FORMATS.includes(value)) fail('INVALID_DOCUMENT_EDIT', `${field} must be one of ${LIST_FORMATS.join(', ')}.`, index, field);
  return value;
}
function positiveInteger(value, index, field) {
  if (!Number.isInteger(value) || value < 1) fail('INVALID_DOCUMENT_EDIT', `${field} must be a positive integer.`, index, field);
  return value;
}

const MAX_INSERTED_LINES = 40;
// Insertions are sent with structuredContent: false, so the library keeps the text literal.
// Each line of the text becomes its own new paragraph (a poem, an address, several items).
function insertedLines(value, index, field) {
  const lines = value.split(/\r\n|\r|\n/);
  if (lines.length > MAX_INSERTED_LINES) fail('UNSUPPORTED_DOCUMENT_EDIT', `One insertion can add at most ${MAX_INSERTED_LINES} paragraphs.`, index, field);
  if (lines.some(line => !line.trim() || line !== line.trim())) {
    fail('UNSUPPORTED_DOCUMENT_EDIT', 'Each line of inserted text becomes its own paragraph and must be non-empty without outer whitespace. Leave out blank lines.', index, field);
  }
  return lines;
}

function validateEdit(edit, index) {
  if (!record(edit) || typeof edit.type !== 'string' || !FIELDS[edit.type]) {
    fail('UNSUPPORTED_DOCUMENT_EDIT', 'This edit action is unsupported. Use a supported text, paragraph, list, comment, or basic-formatting action.', index, 'type');
  }
  // A list item is always inserted after its target, so an explicit "after" is accepted (KB-12).
  if (edit.type === 'insert_list_item' && edit.position !== undefined) {
    if (edit.position !== 'after') fail('UNSUPPORTED_DOCUMENT_EDIT', 'insert_list_item adds the item after its target; to add an item before it, target the previous list item.', index, 'position');
    edit = Object.fromEntries(Object.entries(edit).filter(([key]) => key !== 'position'));
  }
  const allowed = new Set(['type', ...FIELDS[edit.type]]);
  const extra = Object.keys(edit).find(key => !allowed.has(key));
  if (extra) fail('INVALID_DOCUMENT_EDIT', `Field "${extra}" is not supported for ${edit.type}. Its fields are: ${FIELDS[edit.type].join(', ')}.`, index, extra);
  const value = { ...edit, target: text(edit.target, index, 'target') };

  if (edit.type === 'replace_text') {
    value.find = text(edit.find, index, 'find');
    value.replace = text(edit.replace, index, 'replace', true);
    if (/[\r\n]/.test(value.find + value.replace)) fail('UNSUPPORTED_DOCUMENT_EDIT', 'Literal replacement cannot span or insert paragraph breaks.', index, 'find');
    if (value.find === value.replace) fail('UNSUPPORTED_LITERAL_NOOP', 'A literal replacement must change the matched text. Use rewrite_paragraph with the complete desired paragraph to confirm an already-satisfied state.', index, 'replace');
  } else if (edit.type === 'rewrite_paragraph') {
    value.text = text(edit.text, index, 'text');
    if (/[\r\n]/.test(value.text)) fail('UNSUPPORTED_DOCUMENT_EDIT', 'A paragraph rewrite must contain one paragraph without line breaks. For new multi-line text, use insert_paragraph or append_paragraph, where each line becomes its own paragraph. If the paragraph already contains line breaks (Shift+Enter), change its wording with replace_text instead.', index, 'text');
  } else if (edit.type === 'insert_paragraph' || edit.type === 'append_paragraph' || edit.type === 'insert_list_item') {
    value.text = text(edit.text, index, 'text');
    value.lines = insertedLines(value.text, index, 'text');
    value.text = value.lines.join('\n');
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
  } else if (edit.type === 'convert_to_list') {
    if (edit.through !== undefined) {
      value.through = text(edit.through, index, 'through');
      if (value.through === value.target) fail('INVALID_DOCUMENT_EDIT', 'through must be a later paragraph than target; omit it to convert a single paragraph.', index, 'through');
    }
    if (edit.markers !== undefined && !['typed', 'none'].includes(edit.markers)) fail('INVALID_DOCUMENT_EDIT', 'markers must be "typed" or "none".', index, 'markers');
    value.markers = edit.markers ?? 'typed';
    if (edit.format !== undefined) listFormat(edit.format, index);
    if (value.markers === 'none' && edit.format === undefined) fail('INVALID_DOCUMENT_EDIT', 'format is required when markers is "none".', index, 'format');
    if (edit.start !== undefined) {
      if (value.markers !== 'none') fail('INVALID_DOCUMENT_EDIT', 'start is valid only with markers "none"; typed markers supply their own start value.', index, 'start');
      positiveInteger(edit.start, index, 'start');
    }
  } else if (edit.type === 'change_list_format') {
    if (edit.format === undefined) fail('INVALID_DOCUMENT_EDIT', 'format is required.', index, 'format');
    listFormat(edit.format, index);
    if (edit.level !== undefined && (!Number.isInteger(edit.level) || edit.level < 0 || edit.level > 8)) fail('INVALID_DOCUMENT_EDIT', 'level must be an integer from 0 through 8.', index, 'level');
    if (edit.start !== undefined) positiveInteger(edit.start, index, 'start');
  } else if (edit.type === 'continue_list') {
    value.from = text(edit.from, index, 'from');
    if (value.from === value.target) fail('INVALID_DOCUMENT_EDIT', 'from must be an item of an earlier list, not the target itself.', index, 'from');
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
function checkedDescriptor(target, index, expectedKind, field = 'target') {
  if (target?.kind !== expectedKind) fail('INVALID_DOCUMENT_EDIT_TARGET', `This action requires an inspected ${expectedKind === 'document-end' ? 'document-end' : 'paragraph'} handle.`, index, field, true);
  // Resolved handles carry the library's own inspection descriptor, which is always strong.
  return target.descriptor;
}
function checkedInsertionTarget(descriptor, index) {
  if (descriptor.inTable || descriptor.table || descriptor.list || descriptor.headingLevel
    || (descriptor.styleId && /^heading\b/i.test(String(descriptor.styleId)))) {
    // New paragraphs take the target's paragraph style, so the target must be a plain body paragraph.
    fail('UNSUPPORTED_DOCUMENT_EDIT_TARGET', 'The target of a plain paragraph insertion must be a plain body paragraph, because new paragraphs take its style; a heading, list item or table cell cannot be the target. Retarget the same place from a plain neighbour: to add text just before a heading, insert after the plain paragraph above it; to add text under a heading, insert before the plain paragraph below it. Use insert_list_item to add list items. If there is no plain neighbour, say so.', index, 'target', true);
  }
  // Insertion joins the target and new text with a paragraph break, so soft breaks are ambiguous.
  if (descriptor.exactText !== descriptor.exactText.trim() || /[\r\n]/.test(descriptor.exactText)) {
    fail('UNSUPPORTED_DOCUMENT_EDIT_TARGET', 'This target has outer spacing or soft breaks that plain insertion cannot preserve.', index, 'target');
  }
}
function checkedListInsertionTarget(descriptor, index) {
  // The library gives the new item the target's list and level; Word evidence covers bullet/decimal levels 0-1.
  const list = descriptor.list;
  if (descriptor.inTable || descriptor.table || !record(list) || !['bullet', 'decimal'].includes(list.format)
    || ![0, 1].includes(list.level)
    || descriptor.headingLevel || (descriptor.styleId && /^heading\b/i.test(String(descriptor.styleId)))
    || !descriptor.exactText || descriptor.exactText !== descriptor.exactText.trim() || /[\r\n]/.test(descriptor.exactText)) {
    fail('UNSUPPORTED_DOCUMENT_EDIT_TARGET', 'List insertion is supported only beside a non-empty level 0 or 1 bullet/decimal list item without soft breaks, outside tables and headings.', index, 'target');
  }
  return list;
}
function checkedPlainListCandidate(descriptor, index, field) {
  if (descriptor.inTable || descriptor.table || descriptor.list) {
    fail('UNSUPPORTED_DOCUMENT_EDIT_TARGET', 'convert_to_list requires plain non-list paragraphs outside tables. Existing list items change with change_list_format or continue_list.', index, field);
  }
  softBreakFreeListCandidate(descriptor.exactText, index, field);
}
// Soft line breaks inside one paragraph ("A. ...", "B. ..." on separate lines) would become
// one list item holding every line, so convert_to_list refuses them (KB-11).
function softBreakFreeListCandidate(text, index, field) {
  if (/[\r\n]/.test(String(text || ''))) {
    fail('UNSUPPORTED_DOCUMENT_EDIT_TARGET', 'convert_to_list needs one item per paragraph; this paragraph has line breaks inside it (Shift+Enter), so it would become a single list item. Split it into separate paragraphs first.', index, field);
  }
}
function checkedListItem(descriptor, index, field) {
  if (descriptor.inTable || descriptor.table || !record(descriptor.list)) {
    fail('UNSUPPORTED_DOCUMENT_EDIT_TARGET', 'This action requires an inspected list item outside tables. Plain paragraphs convert with convert_to_list.', index, field);
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
  const groupFor = handle => {
    const key = identity(targets.get(handle).descriptor);
    const group = groups.get(key) || { text: [], comments: [], formats: [], lists: [] };
    groups.set(key, group);
    return group;
  };
  for (const item of items) {
    if (LIST_ACTIONS.has(item.edit.type)) {
      groupFor(item.edit.target).lists.push(item);
      if (item.edit.through) groupFor(item.edit.through).lists.push(item);
      continue;
    }
    const group = groupFor(item.edit.target);
    if (TEXT_ACTIONS.has(item.edit.type)) group.text.push(item);
    else if (item.edit.type === 'comment') group.comments.push(item);
    else group.formats.push(item);
  }
  for (const group of groups.values()) {
    if (group.lists.length) {
      const others = [...group.text, ...group.comments, ...group.formats];
      if (group.lists.length > 1) fail('UNSUPPORTED_DOCUMENT_EDIT_COMBINATION', 'A source paragraph can take only one list action per batch.', group.lists[1].actionIndex, 'target');
      const [list] = group.lists;
      const pair = list.edit.type === 'convert_to_list' && !list.edit.through
        && others.length === 1 && others[0].edit.type === 'rewrite_paragraph';
      if (others.length && !pair) {
        fail('UNSUPPORTED_DOCUMENT_EDIT_COMBINATION', 'A list action can share a source paragraph only as one rewrite_paragraph plus one single-paragraph convert_to_list. Apply other changes in a separate stage after fresh inspection.', others[0].actionIndex, 'target');
      }
    }
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
  if (edits.length > 40) refuse('INVALID_DOCUMENT_EDIT_BATCH', 'A document edit batch may contain at most 40 edits.');
  if (!session || typeof session.resolveHandles !== 'function') refuse('INVALID_DOCUMENT_EDIT_SESSION', 'A document inspection session is required.');
  const items = edits.map(validateEdit);
  const handles = [...new Set(items.flatMap(({ edit }) => [edit.target, edit.through, edit.from].filter(Boolean)))];
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
    if (edit.type === 'insert_list_item') checkedListInsertionTarget(descriptor, actionIndex);
    if (edit.type === 'delete_paragraph' && (descriptor.inTable || descriptor.table || descriptor.list)) {
      fail('UNSUPPORTED_DOCUMENT_EDIT_TARGET', 'Deleting a paragraph is supported only for plain non-list paragraphs outside tables.', actionIndex, 'target');
    }
    if (edit.type === 'insert_paragraph' && !descriptor.exactText) fail('UNSUPPORTED_DOCUMENT_EDIT_TARGET', 'Insert before or after a blank paragraph is ambiguous; choose a non-empty paragraph.', actionIndex, 'target');
    if (LIST_ACTIONS.has(edit.type)) {
      const checkRelated = (handle, field, validate) => {
        const related = targets.get(handle);
        const relatedDescriptor = checkedDescriptor(related, actionIndex, 'paragraph', field);
        if (related.fullTextInspected !== true) fail('TARGET_TEXT_NOT_FULLY_INSPECTED', 'This paragraph was only partially shown. Inspect the complete paragraph before editing around it.', actionIndex, field, true);
        validate(relatedDescriptor, actionIndex, field);
        return relatedDescriptor;
      };
      if (edit.type === 'convert_to_list') {
        checkedPlainListCandidate(descriptor, actionIndex, 'target');
        if (edit.through) {
          const through = checkRelated(edit.through, 'through', checkedPlainListCandidate);
          if (through.index <= descriptor.index) fail('INVALID_DOCUMENT_EDIT', 'through must be a later paragraph than target.', actionIndex, 'through');
          const between = result?.source?.parts && inspectDocumentParts(result.source.parts, {
            revisionView: 'accepted', range: { start: descriptor.index, end: through.index }
          });
          for (const paragraph of between?.paragraphs || []) softBreakFreeListCandidate(paragraph.exactText, actionIndex, 'through');
        }
      } else {
        checkedListItem(descriptor, actionIndex, 'target');
        if (edit.type === 'continue_list') checkRelated(edit.from, 'from', checkedListItem);
      }
    }
    if (edit.type === 'comment') {
      const anchor = edit.anchor ?? descriptor.exactText;
      if (!anchor || !descriptor.exactText.includes(anchor)) fail('COMMENT_ANCHOR_REQUIRES_REINSPECTION', 'The comment anchor is not present in inspected text. Inspect the resulting paragraph before commenting.', actionIndex, 'anchor', true);
    }
  }
  combinations(items, targets);

  const operations = [], operationActions = [], expectations = [], locations = [];
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
      case 'insert_list_item': {
        // A plain-text redline adds the new paragraph as a sibling in the target's list and level.
        operation = { type: 'redline', target, modified: `${descriptor.exactText}\n${edit.text}`, structuredContent: false };
        expected.expectedParagraphs = [descriptor.exactText, ...edit.lines];
        expected.expectedListInsertion = { format: descriptor.list.format, level: descriptor.list.level };
        break;
      }
      case 'insert_paragraph': {
        const current = descriptor.exactText;
        operation = { type: 'redline', target, modified: edit.position === 'before' ? `${edit.text}\n${current}` : `${current}\n${edit.text}`, structuredContent: false };
        expected.expectedParagraphs = edit.position === 'before' ? [...edit.lines, current] : [current, ...edit.lines];
        break;
      }
      case 'append_paragraph': {
        const current = descriptor.exactText;
        operation = { type: 'redline', target, modified: current ? `${current}\n${edit.text}` : edit.text, structuredContent: false };
        expected.expectedParagraphs = current ? [current, ...edit.lines] : [...edit.lines];
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
      case 'convert_to_list': {
        const list = { mode: 'convert-plain', strip: edit.markers === 'none' ? 'none' : 'marker',
          ...(edit.format ? { format: edit.format } : {}), ...(edit.start ? { start: edit.start } : {}) };
        operation = { type: 'list-format', target,
          ...(edit.through ? { targetEnd: targetDescriptor(targets.get(edit.through).descriptor) } : {}), list };
        break;
      }
      case 'change_list_format':
        operation = { type: 'list-format', target, list: { mode: 'change-format', format: edit.format, scope: 'list',
          ...(edit.level !== undefined ? { level: edit.level } : {}), ...(edit.start ? { start: edit.start } : {}) } };
        break;
      case 'continue_list':
        operation = { type: 'list-format', target, list: { mode: 'join', continueFrom: targetDescriptor(targets.get(edit.from).descriptor), scope: 'list' } };
        break;
      default: fail('UNSUPPORTED_DOCUMENT_EDIT', 'This edit action is unsupported.', actionIndex, 'type');
    }
    if (operation) { operations.push(operation); operationActions.push(actionIndex); }
    expectations.push(expected);
    locations.push({ handle: edit.target, location: location(resolved) });
  }
  return { operations, operationActions, expectations, locations, requestedCount: edits.length };
}
