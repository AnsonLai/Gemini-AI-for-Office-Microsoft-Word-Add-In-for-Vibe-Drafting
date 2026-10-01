import {
  buildListMarkdown,
  normalizeListItemsWithLevels
} from '@ansonlai/docx-redline-js';
import { validateListRequest } from './agentic-request-validation.js';
import { resolveInsertListItemLevel } from './list-level-utils.js';

const NUMBERING_STYLES = new Set(['decimal', 'lowerAlpha', 'upperAlpha', 'lowerRoman', 'upperRoman']);

function refuse(code, message) {
  const error = new Error(message);
  error.code = code;
  throw error;
}

function sourceParagraphs(source) {
  if (!source || !Array.isArray(source.paragraphs)) {
    refuse('INVALID_LIST_SOURCE', 'The source snapshot has no inspected paragraph list.');
  }
  return source.paragraphs;
}

function paragraphAt(source, oneBasedIndex) {
  const paragraphs = sourceParagraphs(source);
  if (!Number.isInteger(oneBasedIndex) || oneBasedIndex < 1 || oneBasedIndex > paragraphs.length) {
    refuse('INVALID_LIST_TARGET', `Paragraph index ${String(oneBasedIndex)} is outside the inspected source (1-${paragraphs.length}).`);
  }
  const paragraph = paragraphs[oneBasedIndex - 1];
  if (!paragraph || paragraph.index !== oneBasedIndex) {
    refuse('INVALID_LIST_SOURCE', `The source paragraph at index ${oneBasedIndex} is missing or out of order.`);
  }
  return paragraph;
}

function exactText(paragraph) {
  return typeof paragraph.exactText === 'string' ? paragraph.exactText : String(paragraph.text ?? '');
}

function targetDescriptor(paragraph) {
  return {
    index: paragraph.index,
    exactText: exactText(paragraph),
    ...(paragraph.paragraphId ? { paragraphId: paragraph.paragraphId } : {}),
    ...(paragraph.fingerprint ? { fingerprint: paragraph.fingerprint } : {}),
    ...(typeof paragraph.inTable === 'boolean' ? { inTable: paragraph.inTable } : {})
  };
}

function redline(paragraph, modified, targetEnd = null) {
  return {
    type: 'redline',
    target: targetDescriptor(paragraph),
    ...(targetEnd ? { targetEnd: targetDescriptor(targetEnd) } : {}),
    modified
  };
}

function requireSingleLineText(value, label) {
  if (typeof value !== 'string' || !value.trim()) {
    refuse('INVALID_LIST_OPERATION', `${label} must be a non-empty string.`);
  }
  if (/\r|\n/.test(value)) {
    refuse('UNSUPPORTED_LIST_OPERATION', `${label} must contain one paragraph of text.`);
  }
  return value;
}

function compositeMarkerForLevel(level) {
  // 0 is represented by a single-level marker; levels 1–8 use the library's
  // explicit composite-decimal outline marker syntax.
  return level === 0 ? '1.' : `${Array(level + 1).fill('1').join('.')}.`;
}

function numberedLineAtLevel(level, text) {
  return `${compositeMarkerForLevel(level)} ${text}`;
}

function insertListItem(source, request) {
  const target = paragraphAt(source, request.afterParagraphIndex);
  const text = requireSingleLineText(request.text, 'List item text');
  const currentText = exactText(target);
  if (!currentText.trim()) {
    refuse('UNSUPPORTED_LIST_OPERATION', 'Cannot insert a list item after an empty paragraph.');
  }

  if (!target.list?.numId) {
    // The existing tool inserts a plain paragraph when its anchor is not in a
    // list. A localized structural redline keeps the source paragraph intact.
    return [redline(target, `${currentText}\n${text}`)];
  }

  const baseLevel = Number.isInteger(target.list.level) ? target.list.level : 0;
  const level = resolveInsertListItemLevel(baseLevel, request.indentLevel ?? 0).newIlvl;
  const modified = `${currentText}\n${numberedLineAtLevel(level, text)}`;

  if (level > 0 || baseLevel === 0) {
    return [redline(target, modified)];
  }

  // The list insertion heuristic can express explicit levels 1–8. Level 0 is
  // implicit in its marker grammar, so use the supported explicit-range
  // insertion path when a following level-0 sibling can provide that context.
  const paragraphs = sourceParagraphs(source);
  const endOffset = paragraphs.slice(request.afterParagraphIndex).findIndex(paragraph => (
    paragraph.list?.numId === target.list.numId && paragraph.list.level === 0
  ));
  if (endOffset < 0) {
    refuse(
      'UNSUPPORTED_LIST_LEVEL_MAPPING',
      'This library version cannot safely encode an outdent to level 0 without a following level-0 item in the same list.'
    );
  }

  const endIndex = request.afterParagraphIndex + endOffset + 1;
  const range = paragraphs.slice(request.afterParagraphIndex - 1, endIndex);
  if (range.some(paragraph => paragraph.list?.numId !== target.list.numId || !exactText(paragraph).trim())) {
    refuse(
      'UNSUPPORTED_LIST_LEVEL_MAPPING',
      'A same-list, non-empty paragraph range is required to preserve numbering while outdenting to level 0.'
    );
  }

  const rangeLines = [];
  for (const paragraph of range) {
    const originalLevel = Number.isInteger(paragraph.list.level) ? paragraph.list.level : 0;
    rangeLines.push(numberedLineAtLevel(originalLevel, exactText(paragraph)));
    if (paragraph.index === target.index) rangeLines.push(numberedLineAtLevel(level, text));
  }
  return [redline(target, rangeLines.join('\n'), range[range.length - 1])];
}

function editList(source, request) {
  const startIndex = request.startParagraphIndex ?? request.startIndex;
  const endIndex = request.endParagraphIndex ?? request.endIndex;
  const start = paragraphAt(source, startIndex);
  const end = paragraphAt(source, endIndex);
  if (startIndex > endIndex) {
    refuse('INVALID_LIST_TARGET', 'The list range start must be at or before its end.');
  }
  if (!Array.isArray(request.newItems) || request.newItems.length === 0) {
    refuse('INVALID_LIST_OPERATION', 'At least one list item is required.');
  }
  for (const item of request.newItems) requireSingleLineText(String(item ?? ''), 'List item text');

  const listType = request.listType === 'bullet' ? 'bullet' : 'numbered';
  const numberingStyle = NUMBERING_STYLES.has(request.numberingStyle) ? request.numberingStyle : 'decimal';
  const itemsWithLevels = normalizeListItemsWithLevels(request.newItems, { indentSpaces: 4 });
  const listMarkdown = buildListMarkdown(itemsWithLevels, listType, numberingStyle);

  return [redline(start, listMarkdown, end)];
}

function stripManualHeaderNumbering(text) {
  return String(text || '')
    .replace(/^\s*(?:(?:\d+|[a-zA-Z]+|[ivxlcIVXLC]+)[.)]\s*)+/, '')
    .trim();
}

function hasManualHeaderNumbering(text) {
  return /^\s*(?:(?:\d+(?:\.\d+)*\.?|\([\dA-Za-zivxlcIVXLC]+\)|[A-Za-z]\.)\s+)/.test(String(text || ''));
}

function normalizeHeaderText(text) {
  return String(text || '').replace(/\s+/g, ' ').trim();
}

function alphaSequence(index, upper) {
  let number = index;
  let value = '';
  while (number > 0) {
    number -= 1;
    value = String.fromCharCode((upper ? 65 : 97) + (number % 26)) + value;
    number = Math.floor(number / 26);
  }
  return value;
}

function romanSequence(value) {
  const pairs = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let remaining = value;
  let output = '';
  for (const [amount, symbol] of pairs) {
    while (remaining >= amount) {
      output += symbol;
      remaining -= amount;
    }
  }
  return output;
}

function markerForHeader(index, format) {
  switch (format) {
    case 'lowerAlpha': return `${alphaSequence(index, false)}.`;
    case 'upperAlpha': return `${alphaSequence(index, true)}.`;
    case 'lowerRoman': return `${romanSequence(index)}.`;
    case 'upperRoman': return `${romanSequence(index).toUpperCase()}.`;
    default: return `${index}.`;
  }
}

function normalizeHeaderNumberingFormat(format) {
  const mapping = {
    arabic: 'decimal',
    lowerLetter: 'lowerAlpha',
    upperLetter: 'upperAlpha',
    lowerRoman: 'lowerRoman',
    upperRoman: 'upperRoman'
  };
  if (format == null) return 'decimal';
  if (mapping[format]) return mapping[format];
  if (NUMBERING_STYLES.has(format)) return format;
  refuse('INVALID_LIST_REQUEST', `Unsupported numberingFormat: ${String(format)}.`);
}

function convertHeadersToList(source, request) {
  const records = Array.isArray(request.headerRecords)
    ? request.headerRecords.map(record => ({ index: record.paragraphIndex, text: record.text }))
    : (Array.isArray(request.paragraphIndices)
      ? request.paragraphIndices.map((index, inputIndex) => ({ index, text: request.newHeaderTexts?.[inputIndex] }))
      : []);
  if (records.length === 0) {
    refuse('INVALID_LIST_OPERATION', 'At least one header paragraph index is required.');
  }
  if (request.newHeaderTexts != null && !Array.isArray(request.newHeaderTexts)) {
    refuse('INVALID_LIST_OPERATION', 'newHeaderTexts must be an array when provided.');
  }

  const mapped = new Map();
  records.forEach(({ index, text: suppliedText }) => {
    const paragraph = paragraphAt(source, index);
    const headerText = suppliedText == null
      ? stripManualHeaderNumbering(exactText(paragraph))
      : requireSingleLineText(String(suppliedText), 'Header text');
    if (!headerText.trim()) refuse('INVALID_LIST_OPERATION', `Header paragraph ${index} has no text to convert.`);

    const existing = mapped.get(index);
    if (existing && existing.headerText !== headerText) {
      refuse('AMBIGUOUS_LIST_TARGET', `Paragraph ${index} was requested with conflicting header text.`);
    }
    mapped.set(index, { paragraph, headerText });
  });

  const format = normalizeHeaderNumberingFormat(request.numberingFormat);
  const ordered = [...mapped.values()].sort((a, b) => a.paragraph.index - b.paragraph.index);
  return ordered.map(({ paragraph, headerText }, index) => {
    if (paragraph.list?.numId) {
      refuse(
        'UNSUPPORTED_LIST_CONVERSION',
        `Paragraph ${paragraph.index} is already list-bound; the supported standalone conversion path only handles plain header paragraphs.`
      );
    }
    const originalText = exactText(paragraph);
    if (
      !hasManualHeaderNumbering(originalText)
      || normalizeHeaderText(stripManualHeaderNumbering(originalText)) !== normalizeHeaderText(headerText)
    ) {
      refuse(
        'UNSUPPORTED_LIST_CONVERSION',
        `Paragraph ${paragraph.index} needs an existing manual list marker and unchanged header text for the supported standalone conversion path.`
      );
    }
    const modified = `${markerForHeader(index + 1, format)} ${headerText}`;
    return redline(paragraph, modified);
  });
}

/**
 * Converts a list-tool request into canonical 0.8.2 redline operations using
 * one immutable OOXML inspection snapshot. The library owns all mutations and
 * numbering artifacts; this module only maps tool intent and source identities.
 *
 * @param {{documentXml?:string, paragraphs:Array<object>, numberingXml?:string|null, stylesXml?:string|null}} source
 * @param {object} request
 * @returns {Array<object>}
 */
export function planAgenticListOperations(source, request) {
  const paragraphs = sourceParagraphs(source);
  if (!request || typeof request !== 'object' || Array.isArray(request)) {
    refuse('INVALID_LIST_OPERATION', 'A list tool request object is required.');
  }
  const validation = validateListRequest(request, paragraphs.length);
  if (!validation.valid) refuse(validation.error.code, validation.error.message);
  const normalized = validation.request;
  switch (normalized.kind) {
    case 'insert_list_item': return insertListItem(source, normalized);
    case 'edit_list': return editList(source, normalized);
    case 'convert_headers_to_list': return convertHeadersToList(source, normalized);
    default: refuse('UNSUPPORTED_LIST_OPERATION', `Unsupported list tool: ${String(request.tool ?? request.kind)}.`);
  }
}
