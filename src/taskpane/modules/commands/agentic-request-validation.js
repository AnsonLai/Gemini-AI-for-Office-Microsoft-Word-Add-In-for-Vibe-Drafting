const LIST_TYPES = new Set(['bullet', 'numbered']);
const NUMBERING_STYLES = new Set(['decimal', 'lowerAlpha', 'upperAlpha', 'lowerRoman', 'upperRoman']);
const HEADER_NUMBERING_FORMATS = new Set(['arabic', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman']);

function invalid(message) {
  return {
    valid: false,
    error: { code: 'INVALID_LIST_REQUEST', message }
  };
}

function hasOwn(value, key) {
  return Object.prototype.hasOwnProperty.call(value, key);
}

function isPlainRecord(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function validParagraphCount(paragraphCount) {
  return paragraphCount == null
    || (Number.isInteger(paragraphCount) && paragraphCount >= 0);
}

function validatePositiveIndex(value, field, paragraphCount) {
  if (!Number.isInteger(value) || value < 1) {
    return `${field} must be a positive 1-based integer.`;
  }
  if (Number.isInteger(paragraphCount) && value > paragraphCount) {
    return `${field} P${value} is outside the document range 1-${paragraphCount}.`;
  }
  return null;
}

function validateNonEmptyStrings(values, field) {
  if (!Array.isArray(values) || values.length === 0) {
    return `${field} must be a non-empty array of strings.`;
  }
  for (let index = 0; index < values.length; index += 1) {
    if (typeof values[index] !== 'string' || values[index].trim().length === 0) {
      return `${field}[${index}] must be a non-empty string.`;
    }
  }
  return null;
}

function validateEditList(request, paragraphCount) {
  const { startParagraphIndex, endParagraphIndex } = request;
  const startError = validatePositiveIndex(startParagraphIndex, 'startParagraphIndex', paragraphCount);
  if (startError) return invalid(startError);
  const endError = validatePositiveIndex(endParagraphIndex, 'endParagraphIndex', paragraphCount);
  if (endError) return invalid(endError);
  if (endParagraphIndex < startParagraphIndex) {
    return invalid('endParagraphIndex must be greater than or equal to startParagraphIndex.');
  }

  const itemsError = validateNonEmptyStrings(request.newItems, 'newItems');
  if (itemsError) return invalid(itemsError);

  if (typeof request.listType !== 'string' || !LIST_TYPES.has(request.listType)) {
    return invalid('listType must be exactly "bullet" or "numbered".');
  }

  let numberingStyle;
  if (hasOwn(request, 'numberingStyle')) {
    if (typeof request.numberingStyle !== 'string' || !NUMBERING_STYLES.has(request.numberingStyle)) {
      return invalid(`numberingStyle must be one of: ${[...NUMBERING_STYLES].join(', ')}.`);
    }
    numberingStyle = request.numberingStyle;
  } else {
    numberingStyle = 'decimal';
  }

  return {
    valid: true,
    request: {
      kind: 'edit_list',
      startParagraphIndex,
      endParagraphIndex,
      // Preserve indentation and the caller's exact item strings; the existing
      // list-level normalizer owns interpretation of custom indentation.
      newItems: [...request.newItems],
      listType: request.listType,
      numberingStyle
    }
  };
}

function validateInsertListItem(request, paragraphCount) {
  const { afterParagraphIndex, text } = request;
  const indexError = validatePositiveIndex(afterParagraphIndex, 'afterParagraphIndex', paragraphCount);
  if (indexError) return invalid(indexError);
  if (typeof text !== 'string' || text.trim().length === 0) {
    return invalid('text must be a non-empty string.');
  }
  if (/[\r\n\u000b\u000c\u2028\u2029]/.test(text)) {
    return invalid('text must contain a single paragraph without line breaks.');
  }

  let indentLevel = 0;
  if (hasOwn(request, 'indentLevel')) {
    if (!Number.isInteger(request.indentLevel) || ![-1, 0, 1].includes(request.indentLevel)) {
      return invalid('indentLevel must be the integer -1, 0, or 1.');
    }
    indentLevel = request.indentLevel;
  }

  return {
    valid: true,
    request: {
      kind: 'insert_list_item',
      afterParagraphIndex,
      text,
      indentLevel
    }
  };
}

function validateConvertHeadersToList(request, paragraphCount) {
  const indices = request.paragraphIndices;
  if (!Array.isArray(indices) || indices.length === 0) {
    return invalid('paragraphIndices must be a non-empty array of positive 1-based integers.');
  }

  for (let index = 0; index < indices.length; index += 1) {
    const indexError = validatePositiveIndex(indices[index], `paragraphIndices[${index}]`, paragraphCount);
    if (indexError) return invalid(indexError);
  }

  const hasHeaderTexts = hasOwn(request, 'newHeaderTexts');
  if (hasHeaderTexts) {
    const textError = validateNonEmptyStrings(request.newHeaderTexts, 'newHeaderTexts');
    if (textError) return invalid(textError);
    if (request.newHeaderTexts.length !== indices.length) {
      return invalid('newHeaderTexts must contain exactly one string for each paragraphIndices entry.');
    }
  }

  const seenIndices = new Set();
  for (const paragraphIndex of indices) {
    if (seenIndices.has(paragraphIndex)) {
      return invalid(`paragraphIndices contains duplicate P${paragraphIndex}; each header must have one unambiguous text mapping.`);
    }
    seenIndices.add(paragraphIndex);
  }

  // Bind text to its original index before sorting. This prevents input-order
  // changes from attaching a header's replacement text to another paragraph.
  const headerRecords = indices.map((paragraphIndex, index) => ({
    paragraphIndex,
    ...(hasHeaderTexts ? { text: request.newHeaderTexts[index] } : {})
  })).sort((left, right) => left.paragraphIndex - right.paragraphIndex);

  let numberingFormat;
  if (hasOwn(request, 'numberingFormat')) {
    if (typeof request.numberingFormat !== 'string' || !HEADER_NUMBERING_FORMATS.has(request.numberingFormat)) {
      return invalid(`numberingFormat must be one of: ${[...HEADER_NUMBERING_FORMATS].join(', ')}.`);
    }
    numberingFormat = request.numberingFormat;
  } else {
    numberingFormat = 'arabic';
  }

  return {
    valid: true,
    request: {
      kind: 'convert_headers_to_list',
      paragraphIndices: headerRecords.map(header => header.paragraphIndex),
      ...(hasHeaderTexts ? { newHeaderTexts: headerRecords.map(header => header.text) } : {}),
      numberingFormat,
      headerRecords
    }
  };
}

/**
 * Validate and normalize a list-related tool request without Word/Office state.
 *
 * `paragraphCount` is optional. When supplied it must be a non-negative integer
 * and all 1-based paragraph indexes must fit within that source snapshot.
 *
 * @param {object} request - raw `insert_list_item`, `edit_list`, or `convert_headers_to_list` args
 * @param {number|null} [paragraphCount] - inspected source paragraph count
 * @returns {{ valid: true, request: object }|{ valid: false, error: { code: string, message: string } }}
 */
export function validateListRequest(request, paragraphCount = null) {
  if (!isPlainRecord(request)) {
    return invalid('A list tool request must be an object.');
  }
  if (!validParagraphCount(paragraphCount)) {
    return invalid('paragraphCount must be a non-negative integer when provided.');
  }

  const hasListFields = ['startParagraphIndex', 'endParagraphIndex', 'newItems', 'listType']
    .some(field => hasOwn(request, field));
  const hasHeaderFields = hasOwn(request, 'paragraphIndices');
  const hasInsertFields = hasOwn(request, 'afterParagraphIndex') || hasOwn(request, 'indentLevel');
  const requestKindCount = Number(hasListFields) + Number(hasHeaderFields) + Number(hasInsertFields);
  if (requestKindCount !== 1) {
    return invalid('Request must contain arguments for exactly one of insert_list_item, edit_list, or convert_headers_to_list.');
  }

  if (hasHeaderFields) return validateConvertHeadersToList(request, paragraphCount);
  if (hasInsertFields) return validateInsertListItem(request, paragraphCount);
  return validateEditList(request, paragraphCount);
}
