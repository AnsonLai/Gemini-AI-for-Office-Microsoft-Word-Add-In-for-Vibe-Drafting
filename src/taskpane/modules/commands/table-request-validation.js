const TABLE_ACTIONS = new Set(['replace_content', 'add_row', 'delete_row', 'update_cell']);

function invalid(message) {
  return {
    valid: false,
    error: { code: 'INVALID_TABLE_REQUEST', message }
  };
}

function hasOwn(value, key) {
  return Object.prototype.hasOwnProperty.call(value, key);
}

function isRecord(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function validateCount(value, field, { allowZero = false } = {}) {
  const minimum = allowZero ? 0 : 1;
  if (!Number.isInteger(value) || value < minimum) {
    return `${field} must be an integer greater than or equal to ${minimum}.`;
  }
  return null;
}

function readLiveDimensions(live) {
  if (live == null) return { valid: true, dimensions: {} };
  if (!isRecord(live)) return { valid: false, message: 'Live table dimensions must be an object.' };

  const dimensions = {};
  for (const [field, allowZero] of [['paragraphCount', true], ['rowCount', false], ['columnCount', false]]) {
    if (!hasOwn(live, field)) continue;
    const error = validateCount(live[field], field, { allowZero });
    if (error) return { valid: false, message: error };
    dimensions[field] = live[field];
  }

  if (hasOwn(live, 'rowCellCounts')) {
    if (!Array.isArray(live.rowCellCounts) || live.rowCellCounts.length === 0) {
      return { valid: false, message: 'rowCellCounts must be a non-empty array of positive integers.' };
    }
    for (let index = 0; index < live.rowCellCounts.length; index += 1) {
      const error = validateCount(live.rowCellCounts[index], `rowCellCounts[${index}]`);
      if (error) return { valid: false, message: error };
    }
    if (hasOwn(dimensions, 'rowCount') && live.rowCellCounts.length !== dimensions.rowCount) {
      return { valid: false, message: 'rowCellCounts length must match rowCount.' };
    }
    dimensions.rowCellCounts = [...live.rowCellCounts];
    if (!hasOwn(dimensions, 'rowCount')) dimensions.rowCount = live.rowCellCounts.length;
  }

  return { valid: true, dimensions };
}

function validateCoordinate(value, field, upperBound) {
  if (!Number.isInteger(value) || value < 0) {
    return `${field} must be a non-negative 0-based integer.`;
  }
  if (Number.isInteger(upperBound) && value >= upperBound) {
    return `${field} ${value} is outside the valid range 0-${upperBound - 1}.`;
  }
  return null;
}

function validateContentStrings(values, field) {
  if (!Array.isArray(values) || values.length === 0) {
    return `${field} must be a non-empty array of strings.`;
  }
  for (let index = 0; index < values.length; index += 1) {
    if (typeof values[index] !== 'string') {
      return `${field}[${index}] must be a string.`;
    }
  }
  return null;
}

function normalizeReplaceContent(content, dimensions) {
  if (!Array.isArray(content) || content.length === 0) {
    return { error: 'content for replace_content must be a non-empty 2D array of strings.' };
  }

  const rows = [];
  for (let rowIndex = 0; rowIndex < content.length; rowIndex += 1) {
    const row = content[rowIndex];
    const rowError = validateContentStrings(row, `content[${rowIndex}]`);
    if (rowError) return { error: rowError };

    if (Number.isInteger(dimensions.rowCount) && rowIndex >= dimensions.rowCount) {
      return { error: `content row ${rowIndex} is outside the table row range 0-${dimensions.rowCount - 1}.` };
    }
    const rowWidth = dimensions.rowCellCounts?.[rowIndex] ?? dimensions.columnCount;
    if (Number.isInteger(rowWidth) && row.length > rowWidth) {
      return { error: `content[${rowIndex}] has ${row.length} cells, exceeding that row's ${rowWidth}-cell capacity.` };
    }
    rows.push([...row]);
  }

  return { content: rows };
}

function normalizeAddRow(content, dimensions) {
  let row;
  if (!Array.isArray(content) || content.length === 0) {
    return { error: 'content for add_row must be a non-empty array of cell strings.' };
  }
  if (Array.isArray(content[0])) {
    if (content.length !== 1) {
      return { error: 'add_row accepts exactly one row; extra nested rows would be ignored.' };
    }
    row = content[0];
  } else {
    row = content;
  }

  const rowError = validateContentStrings(row, 'content row');
  if (rowError) return { error: rowError };

  const appendTemplateWidth = dimensions.rowCellCounts?.at(-1) ?? dimensions.columnCount;
  if (Number.isInteger(appendTemplateWidth) && row.length > appendTemplateWidth) {
    return { error: `add_row has ${row.length} cells, exceeding the appended row's ${appendTemplateWidth}-cell capacity.` };
  }
  return { content: [[...row]] };
}

function normalizeUpdateCell(content) {
  let value;
  if (typeof content === 'string') {
    value = content;
  } else if (Array.isArray(content)) {
    if (content.length !== 1) {
      return { error: 'update_cell content must describe exactly one cell.' };
    }
    if (Array.isArray(content[0])) {
      if (content[0].length !== 1) {
        return { error: 'update_cell content must describe exactly one cell.' };
      }
      value = content[0][0];
    } else {
      value = content[0];
    }
  } else {
    return { error: 'content for update_cell must be a string or a one-cell array.' };
  }

  if (typeof value !== 'string') return { error: 'update_cell value must be a string.' };
  return { content: [[value]] };
}

/**
 * Validate and normalize an edit_table request without Office/Word dependencies.
 *
 * `live` may provide paragraphCount, rowCount, columnCount, and rowCellCounts
 * after the executor has loaded the target document/table. The normalized
 * content is always a 2D string array for actions that use cell values.
 */
export function validateTableRequest(request, live = {}) {
  if (!isRecord(request)) return invalid('The table request must be an object.');

  const paragraphIndexError = validateCount(request.paragraphIndex, 'paragraphIndex');
  if (paragraphIndexError) return invalid(paragraphIndexError.replace('integer greater than or equal to 1', 'positive 1-based integer'));

  const dimensionsResult = readLiveDimensions(live);
  if (!dimensionsResult.valid) return invalid(dimensionsResult.message);
  const dimensions = dimensionsResult.dimensions;
  if (Number.isInteger(dimensions.paragraphCount) && request.paragraphIndex > dimensions.paragraphCount) {
    return invalid(`paragraphIndex P${request.paragraphIndex} is outside the document range 1-${dimensions.paragraphCount}.`);
  }

  if (typeof request.action !== 'string') return invalid('action must be one of: replace_content, add_row, delete_row, update_cell.');
  const action = request.action.trim().toLowerCase();
  if (!TABLE_ACTIONS.has(action)) return invalid('action must be one of: replace_content, add_row, delete_row, update_cell.');

  const normalized = { kind: 'edit_table', paragraphIndex: request.paragraphIndex, action };
  if (action === 'replace_content') {
    if (!hasOwn(request, 'content')) return invalid('replace_content requires content.');
    if (hasOwn(request, 'targetRow') || hasOwn(request, 'targetColumn')) {
      return invalid('replace_content does not use targetRow or targetColumn.');
    }
    const content = normalizeReplaceContent(request.content, dimensions);
    if (content.error) return invalid(content.error);
    normalized.content = content.content;
  } else if (action === 'add_row') {
    if (!hasOwn(request, 'content')) return invalid('add_row requires content.');
    if (hasOwn(request, 'targetRow') || hasOwn(request, 'targetColumn')) {
      return invalid('add_row always appends and does not accept row or column coordinates.');
    }
    const content = normalizeAddRow(request.content, dimensions);
    if (content.error) return invalid(content.error);
    normalized.content = content.content;
  } else if (action === 'delete_row') {
    if (!hasOwn(request, 'targetRow')) return invalid('delete_row requires targetRow.');
    if (hasOwn(request, 'content') || hasOwn(request, 'targetColumn')) {
      return invalid('delete_row uses targetRow only; content and targetColumn are unsupported.');
    }
    const coordinateError = validateCoordinate(request.targetRow, 'targetRow', dimensions.rowCount);
    if (coordinateError) return invalid(coordinateError);
    normalized.targetRow = request.targetRow;
  } else {
    if (!hasOwn(request, 'targetRow') || !hasOwn(request, 'targetColumn')) {
      return invalid('update_cell requires targetRow and targetColumn.');
    }
    if (!hasOwn(request, 'content')) return invalid('update_cell requires content.');
    const rowError = validateCoordinate(request.targetRow, 'targetRow', dimensions.rowCount);
    if (rowError) return invalid(rowError);
    const rowWidth = dimensions.rowCellCounts?.[request.targetRow] ?? dimensions.columnCount;
    const columnError = validateCoordinate(request.targetColumn, 'targetColumn', rowWidth);
    if (columnError) return invalid(columnError);
    const content = normalizeUpdateCell(request.content);
    if (content.error) return invalid(content.error);
    normalized.targetRow = request.targetRow;
    normalized.targetColumn = request.targetColumn;
    normalized.content = content.content;
  }

  return { valid: true, request: normalized };
}
