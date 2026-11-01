// One authoritative model protocol. Action-specific validation belongs to the intent compiler.
const EDIT_TYPES = ['replace_text', 'rewrite_paragraph', 'delete_paragraph', 'insert_paragraph', 'append_paragraph', 'insert_list_item', 'comment', 'format_text', 'convert_to_list', 'change_list_format', 'continue_list'];
const LIST_FORMATS = ['decimal', 'upperLetter', 'lowerLetter', 'upperRoman', 'lowerRoman', 'bullet'];
const DECLARATIONS = [{ function_declarations: [
  { name: 'inspect_document', description: 'Inspect accepted document text and obtain fresh opaque target handles. Search and surrounding context are library-backed; use cursor to continue a page.', parameters: {
    type: 'OBJECT', properties: {
      search: { type: 'STRING' }, around: { type: 'INTEGER', description: 'Surrounding paragraphs, 0 through 5.' },
      limit: { type: 'INTEGER', description: 'At most 8 matches.' }, cursor: { type: 'STRING' },
      includeDocumentEnd: { type: 'BOOLEAN', description: 'Request an append-only end handle.' },
    },
  } },
  { name: 'apply_document_edits', description: 'Validate and apply every independent requested intent as one atomic batch using previously inspected handles. Unsupported or invalid mixed batches make no changes.', parameters: {
    type: 'OBJECT', required: ['edits'], properties: { edits: { type: 'ARRAY', items: {
      type: 'OBJECT', required: ['type', 'target'], properties: {
        type: { type: 'STRING', enum: EDIT_TYPES, description: 'insert_paragraph, append_paragraph and insert_list_item turn each line of text into its own new paragraph or item; put multi-paragraph text (recitals, clauses, poems) one paragraph per line, never run together or joined with slashes.format_text formats one piece of text; send one per separate piece, several per paragraph if needed, with occurrence for repeated words. insert_list_item adds plain-text items at the same level as a fully inspected level 0/1 bullet or decimal target outside a table. convert_to_list turns plain paragraphs into a new list (target, optional later through for a contiguous run). change_list_format changes the numbering format of the whole list of the inspected list item. continue_list makes the target list continue the earlier list that from belongs to.' }, target: { type: 'STRING', description: 'Current opaque inspected paragraph handle; append requires the document-end handle.' },
        find: { type: 'STRING' }, replace: { type: 'STRING' }, text: { type: 'STRING' },
        occurrence: { type: 'INTEGER', description: '1-based exact text occurrence.' },
        position: { type: 'STRING', enum: ['before', 'after'] }, anchor: { type: 'STRING' },
        through: { type: 'STRING', description: 'convert_to_list only: handle of a later plain paragraph that ends a contiguous run starting at target.' },
        from: { type: 'STRING', description: 'continue_list only: handle of an item of the earlier list to continue.' },
        format: { type: 'STRING', enum: LIST_FORMATS, description: 'convert_to_list: required with markers none, otherwise disambiguates typed markers such as I. or C. change_list_format: the new numbering format.' },
        markers: { type: 'STRING', enum: ['typed', 'none'], description: 'convert_to_list only. typed (default): the paragraph text starts with a typed marker such as 1., A., iv. or - that becomes the list label. none: the text has no marker; requires format.' },
        start: { type: 'INTEGER', description: 'convert_to_list (markers none only) or change_list_format: first number of the list.' },
        level: { type: 'INTEGER', description: 'change_list_format only: list level to redefine, 0 through 8; default is the target level.' },
        properties: { type: 'OBJECT', properties: {
          bold: { type: 'BOOLEAN' }, italic: { type: 'BOOLEAN' }, underline: { type: 'BOOLEAN' },
          strikethrough: { type: 'BOOLEAN' }, highlight: { type: 'STRING', enum: ['yellow', 'none'] },
        } },
      },
    } } },
  } },
] }];
const NAMES = Object.freeze(DECLARATIONS[0].function_declarations.map(tool => tool.name));
const EXECUTORS = { inspect_document: 'inspectDocument', apply_document_edits: 'applyDocumentEdits' };
export const getToolDeclarations = () => JSON.parse(JSON.stringify(DECLARATIONS));
export function createToolRegistry(executors = {}) {
  function describe(name) {
    const known = NAMES.includes(name), mutating = name === 'apply_document_edits';
    return { known, name, mutating,
      instruction: mutating ? 'Apply the requested document batch' : 'Inspect the document',
      progressText: mutating ? 'Validating the document edit batch...' : 'Inspecting the document...',
    };
  }
  function validate(name, args) {
    if (!NAMES.includes(name)) return { valid: false, error: { code: 'UNSUPPORTED_TOOL', message: 'This tool is outside the supported document workflow.' } };
    const allowed = name === 'inspect_document' ? ['search', 'around', 'limit', 'cursor', 'includeDocumentEnd'] : ['edits'];
    if (!args || typeof args !== 'object' || Array.isArray(args) || Object.keys(args).some(key => !allowed.includes(key))) {
      return { valid: false, error: { code: 'INVALID_TOOL_ARGUMENTS', message: 'Use only the declared tool arguments.' } };
    }
    if (name === 'apply_document_edits' && (!Array.isArray(args.edits) || args.edits.length < 1 || args.edits.length > 40)) {
      return { valid: false, error: { code: 'INVALID_DOCUMENT_EDIT_BATCH', message: 'A batch requires 1 through 40 edits.' } };
    }
    if (name === 'inspect_document' && (
      ['search', 'cursor'].some(key => args[key] !== undefined && typeof args[key] !== 'string')
      || ['around', 'limit'].some(key => args[key] !== undefined && !Number.isInteger(args[key]))
      || (args.includeDocumentEnd !== undefined && typeof args.includeDocumentEnd !== 'boolean'))) {
      return { valid: false, error: { code: 'INVALID_TOOL_ARGUMENTS', message: 'Inspection arguments must use their declared types.' } };
    }
    return { valid: true };
  }
  async function execute(name, args = {}) {
    const description = describe(name), validation = validate(name, args);
    let rawResult;
    if (!validation.valid) rawResult = { status: 'error', success: false, written: false, writeAttempted: false, mutationOutcome: 'refused', error: validation.error, message: `No changes made. ${validation.error.message}` };
    else if (typeof executors[EXECUTORS[name]] !== 'function') rawResult = { status: 'error', error: { code: 'TOOL_UNAVAILABLE', message: 'The document tool could not be loaded.' } };
    else rawResult = await executors[EXECUTORS[name]](args);
    const success = description.mutating ? rawResult.success === true && rawResult.completion === true && rawResult.verified === true : rawResult.status === 'ok';
    const hostWriteEvidence = { written: rawResult.written === true, writeAttempted: rawResult.writeAttempted === true, mutationOutcome: rawResult.mutationOutcome || null };
    return {
      name, instruction: description.instruction, mutating: description.mutating, rawResult,
      resultText: JSON.stringify(rawResult), success, recordSuccess: success, recordResult: rawResult,
      mutation: description.mutating ? rawResult : null, hostWriteEvidence,
      requiresInspection: description.mutating && (rawResult.requiresInspection === true || ['indeterminate', 'partial', 'applied_with_host_error'].includes(rawResult.mutationOutcome)
        || (hostWriteEvidence.writeAttempted && !success)),
      completionText: rawResult.message || rawResult.error?.message || 'Inspection complete.',
      completionWithCheckpoint: success && hostWriteEvidence.written && Number.isInteger(rawResult.checkpointIndex),
      checkpointIndex: rawResult.checkpointIndex,
      display: { show: description.mutating || !success },
    };
  }
  return Object.freeze({ recoveryToolNames: NAMES, describe, validate, execute });
}
