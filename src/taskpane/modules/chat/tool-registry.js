const GEMINI_TOOLS = [
      {
        function_declarations: [
          {
            name: "apply_redlines",
            description: "Applies suggested edits to the document. Use this tool whenever the user asks to 'edit text', 'change text', 'modify', 'add', 'delete', 'reword', 'rephrase', 'update', 'bold', 'italicize', 'underline', 'strikethrough', remove such formatting (e.g. 'unbold'), convert text into a table, or apply inline TEXT FORMATTING to existing paragraphs.\n\nIMPORTANT - FORMATTING RULES:\n- Bold: **text**\n- Italic: *text*\n- Underline: ++text++\n- Strikethrough: ~~text~~\n\nIMPORTANT - LIST RULES:\n- Use Markdown syntax for lists. \n- For Bullet Lists: Use '* item'. For nested items, indent with 4 spaces (e.g., '    * sub-item').\n- For Numbered Lists: Use '1. item', 'a. item', 'i. item', etc. explicitly. \n- For Nested Numbering: Use '1.1.', '1.1.1.' styles if appropriate. \n- DO NOT use simple hyphens ('-') if you intend to create a structured or numbered list. \n- INDENTATION is critical for sub-levels. Use 2 or 4 spaces.\n\nIMPORTANT - TABLE RULES:\n- Use apply_redlines for converting normal paragraphs into a new table.\n- In the instruction, explicitly identify the full paragraph range to replace (for example: 'Replace P4 through P6 with a two-column markdown table').\n- Require a complete multiline GitHub Markdown table with a header row, separator row, and data row(s).\n- NEVER ask for a single pipe-delimited line like 'A|B|C'; that is plain text, not a Word table.\n- Preserve multi-line source blocks by using additional table rows. Do not put HTML tags such as <br> inside markdown table cells.\n\nFor full list structure conversions (like turning multiple lines into A., B., C. or 1., 2., 3. list items), prefer the dedicated list tools.\n\nDo NOT suggest changes in the chat; always use this tool to apply them directly. The edits will be applied under track changes (redlines). NEVER say you have applied edits unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The specific instruction for how to edit the document (e.g., 'Change Lessor to Landlord', 'Fix spelling', 'Reword the introduction'). For text-to-table conversions, include the exact source paragraph range and the complete multiline markdown table to insert; do not give only a vague instruction like 'turn these into a table'.",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "insert_comment",
            description: "Inserts comments into the document based on the user's instruction. Use this tool to flag risks, add notes, or review specific sections. NEVER say you have inserted comments unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The instruction for what to comment on and what to say (e.g., 'Flag all risky clauses', 'Comment on the first paragraph').",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "highlight_text",
            description: "Highlights text with a colored background marker (like a highlighter pen). ONLY use this tool when the user EXPLICITLY asks to 'highlight' text. Do NOT use this for formatting requests like 'bold', 'italicize', or general emphasis - those should use apply_redlines with markdown syntax instead. Use this tool ONLY for explicit highlight requests like 'highlight all dates in yellow' or 'mark these terms with highlighting'. NEVER say you have highlighted text unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The instruction for what to highlight (e.g., 'Highlight all dates', 'Mark placeholders').",
                },
                color: {
                  type: "STRING",
                  enum: ["yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "gray25", "gray50", "black", "white"],
                  description: "Optional: highlight color. Default is 'yellow'. Options include: yellow, green, cyan, magenta, blue, red, and dark variants.",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "perform_research",
            description: "Performs a Google Search to answer questions that require external knowledge, facts, or up-to-date information. Use this when the user asks for information not in the document.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The search query to send to Google.",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "navigate_to_section",
            description: "Navigates to and selects a specific section of the document. Use this when the user asks to go to, scroll to, find, or jump to a particular part of the document (e.g., 'go to the introduction', 'scroll to paragraph 5', 'find the signature block', 'show me the definitions section'). This helps users quickly locate relevant content without manually scrolling.",
            parameters: {
              type: "OBJECT",
              properties: {
                instruction: {
                  type: "STRING",
                  description: "The navigation instruction describing what section to go to (e.g., 'go to paragraph 3', 'find the table of contents', 'scroll to the conclusion', 'show me where parties are defined').",
                },
              },
              required: ["instruction"],
            },
          },
          {
            name: "edit_list",
            description: "Edit an entire list as a unit. Use this when you need to modify, add, or remove items from a bulleted or numbered list. This preserves list formatting and structure better than apply_redlines. Look for paragraphs with |ListNumber or |ListBullet in the context. For numbered lists, you can specify different numbering styles: '1, 2, 3' (decimal - default), 'a, b, c' (lowerAlpha), 'A, B, C' (upperAlpha), 'i, ii, iii' (lowerRoman), or 'I, II, III' (upperRoman). NEVER say you have edited a list unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                startParagraphIndex: {
                  type: "INTEGER",
                  description: "The paragraph index of the FIRST item in the list (e.g., 3 for [P3])",
                },
                endParagraphIndex: {
                  type: "INTEGER",
                  description: "The paragraph index of the LAST item in the list",
                },
                newItems: {
                  type: "ARRAY",
                  items: { type: "STRING" },
                  description: "The new list items in order. Each string is one list item text (without bullets/numbers).",
                },
                listType: {
                  type: "STRING",
                  enum: ["bullet", "numbered"],
                  description: "The type of list to create",
                },
                numberingStyle: {
                  type: "STRING",
                  enum: ["decimal", "lowerAlpha", "upperAlpha", "lowerRoman", "upperRoman"],
                  description: "Optional: For numbered lists, the numbering style to use. Default is 'decimal' (1, 2, 3). Options: 'decimal' (1, 2, 3), 'lowerAlpha' (a, b, c), 'upperAlpha' (A, B, C), 'lowerRoman' (i, ii, iii), 'upperRoman' (I, II, III).",
                },
              },
              required: ["startParagraphIndex", "endParagraphIndex", "newItems", "listType"],
            },
          },
          {
            name: "insert_list_item",
            description: "Insert a single list item after a specific paragraph. Use this for surgical additions to an existing list - it inherits the numbering format from the paragraph you insert after. Much better than edit_list when you only need to add one or two items. Do NOT include numbering markers in the text - Word will add them automatically.",
            parameters: {
              type: "OBJECT",
              properties: {
                afterParagraphIndex: {
                  type: "INTEGER",
                  description: "The paragraph index to insert after (e.g., 5 to insert after [P5])",
                },
                text: {
                  type: "STRING",
                  description: "The text content of the new list item (WITHOUT any numbering like '1.' or '1.1.' - Word adds these automatically)",
                },
                indentLevel: {
                  type: "INTEGER",
                  description: "Optional: Relative indentation from the paragraph you're inserting after. Allowed values: -1 (one level shallower), 0 (same level, default), 1 (one level deeper). Values outside -1..1 are treated as invalid and clamped.",
                },
              },
              required: ["afterParagraphIndex", "text"],
            },
          },
          {
            name: "edit_table",
            description: "Edit content or add/remove rows in an existing table. Look for paragraphs with |T:row,col in the context. NEVER say you have edited a table unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                paragraphIndex: {
                  type: "INTEGER",
                  description: "Any paragraph index that is part of the table (has T:row,col marker)",
                },
                action: {
                  type: "STRING",
                  enum: ["replace_content", "add_row", "delete_row", "update_cell"],
                  description: "The table operation to perform",
                },
                content: {
                  type: "ARRAY",
                  items: { type: "ARRAY", items: { type: "STRING" } },
                  description: "For replace_content: rows of cell strings [[row1cells], [row2cells]]. For add_row: exactly one row [[cell1, cell2]]. For update_cell: exactly one cell [[new text]]. Omit for delete_row.",
                },
                targetRow: {
                  type: "INTEGER",
                  description: "For delete_row/update_cell: the 0-based row index. add_row appends at the end.",
                },
                targetColumn: {
                  type: "INTEGER",
                  description: "For update_cell: the 0-based column index",
                },
              },
              required: ["paragraphIndex", "action"],
            },
          },
          {
            name: "edit_section",
            description: "Edit a document section as a unit. Use this for legal contracts where numbered/lettered items serve as section headers (marked with §) followed by body text (marked with §N). This preserves the section structure and list numbering. NEVER say you have edited a section unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                sectionHeaderIndex: {
                  type: "INTEGER",
                  description: "The paragraph index of the section header (the list item marked with §, e.g., '1. Definitions')",
                },
                newHeaderText: {
                  type: "STRING",
                  description: "Optional: new text for the section header. The list number/letter is automatically preserved.",
                },
                newBodyParagraphs: {
                  type: "ARRAY",
                  items: { type: "STRING" },
                  description: "Optional: new body paragraphs for this section. Each string becomes one paragraph. Omit to keep existing body.",
                },
                preserveSubsections: {
                  type: "BOOLEAN",
                  description: "If true, only edits body text until the next subsection. If false or omitted, replaces entire section including subsections.",
                },
              },
              required: ["sectionHeaderIndex"],
            },
          },
          {
            name: "convert_headers_to_list",
            description: "Convert non-contiguous headers to a numbered list. Use this when headers like '1. PURPOSE', '2. DEFINITION', '3. EXCLUSIONS' have body text between them and need to be converted to a proper auto-numbered list. The tool strips manual numbering and creates a Word list where all headers share continuous numbering. Supports different formats: 1,2,3 or a,b,c or i,ii,iii. NEVER say you have converted headers unless you have successfully called this tool.",
            parameters: {
              type: "OBJECT",
              properties: {
                paragraphIndices: {
                  type: "ARRAY",
                  items: { type: "INTEGER" },
                  description: "Array of 1-based paragraph indices of the headers to convert (e.g., [3, 7, 15] for headers at P3, P7, P15)",
                },
                newHeaderTexts: {
                  type: "ARRAY",
                  items: { type: "STRING" },
                  description: "Optional: new text for each header (without numbers). If omitted, existing text is used with manual numbers stripped.",
                },
                numberingFormat: {
                  type: "STRING",
                  enum: ["arabic", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman"],
                  description: "Optional: numbering format. 'arabic' = 1,2,3 (default), 'lowerLetter' = a,b,c, 'upperLetter' = A,B,C, 'lowerRoman' = i,ii,iii, 'upperRoman' = I,II,III",
                },
              },
              required: ["paragraphIndices"],
            },
          },
          {
            name: "get_selection_stats",
            description: "Calculates the exact word count and character count (stripping out all formatting like markdown, internal tags, and HTML) of a selection, text snippet, or specified paragraph range. Use this tool whenever the user asks for word count, character count, or document statistics.",
            parameters: {
              type: "OBJECT",
              properties: {
                text: {
                  type: "STRING",
                  description: "Optional: The text string to count. If user provided a specific quote or text, pass it here.",
                },
                startParagraphIndex: {
                  type: "INTEGER",
                  description: "Optional: The starting 1-based paragraph index (e.g., 1 for [P1])",
                },
                endParagraphIndex: {
                  type: "INTEGER",
                  description: "Optional: The ending 1-based paragraph index (e.g., 3 for [P3])",
                },
              },
            },
          },
        ],
      },
    ];

const TOOL_DEFINITIONS = GEMINI_TOOLS[0].function_declarations;

const COMMON_MESSAGE_MUTATION = {
  mutating: true,
  resultText: result => result?.message,
  success: result => !!result?.success,
  recordSuccess: result => result?.success,
  recordResult: (_result, resultText) => resultText,
  display: (result, resultText) => ({ show: true, text: resultText, withCheckpoint: !!result?.success })
};

const TOOL_METADATA = {
  apply_redlines: {
    executorName: 'executeRedline', mutating: true,
    executorArgs: (args, context) => [args.instruction, context.docText ?? '', context.docSourceBaseline],
    progressText: args => `Applying edits: "${args.instruction}"...`,
    instruction: args => args.instruction,
    resultText: result => result?.message,
    success: result => result?.written === true && result?.status !== 'error',
    recordSuccess: result => result?.written === true && result?.status !== 'error',
    recordResult: (_result, resultText) => resultText,
    display: (result, resultText) => result?.showToUser
      ? { show: true, text: resultText, withCheckpoint: true }
      : { show: false, text: 'Redline tool completed without a confirmed write.', withCheckpoint: false }
  },
  insert_comment: {
    executorName: 'executeComment', mutating: true,
    executorArgs: (args, context) => [args.instruction, context.docText ?? ''],
    progressText: args => `Inserting comments: "${args.instruction}"...`,
    instruction: args => args.instruction,
    resultText: result => result?.message,
    success: result => result?.status !== 'error' && result?.written === true,
    recordSuccess: result => result?.showToUser,
    recordResult: (_result, resultText) => resultText,
    display: (result, resultText) => result?.showToUser
      ? { show: true, text: resultText, withCheckpoint: true }
      : { show: false, text: 'Fallback in progress (0 comments): ' + (resultText || ''), withCheckpoint: false }
  },
  highlight_text: {
    executorName: 'executeHighlight', mutating: true,
    executorArgs: (args, context) => [args.instruction, context.docText ?? '', args.color || 'yellow'],
    progressText: args => `Highlighting text: "${args.instruction}"...`,
    instruction: args => args.instruction,
    resultText: result => result?.message,
    success: result => result?.status !== 'error' && result?.written === true,
    recordSuccess: result => result?.showToUser,
    recordResult: (_result, resultText) => resultText,
    display: (result, resultText) => result?.showToUser
      ? { show: true, text: resultText, withCheckpoint: true }
      : { show: false, text: 'Fallback in progress (0 highlights): ' + (resultText || ''), withCheckpoint: false }
  },
  perform_research: {
    executorName: 'executeResearch', mutating: false,
    executorArgs: args => [args.instruction],
    progressText: args => `Researching: "${args.instruction}"...`,
    instruction: args => args.instruction,
    resultText: result => result,
    success: () => true,
    recordSuccess: () => true,
    recordResult: (_result, resultText) => resultText,
    display: (_result, _resultText, args) => ({ show: true, text: `Found search results for: "${args.instruction}"`, withCheckpoint: false })
  },
  navigate_to_section: {
    executorName: 'executeNavigate', mutating: false,
    executorArgs: (args, context) => [args.instruction, context.docText ?? ''],
    progressText: args => `Navigating to: "${args.instruction}"...`,
    instruction: args => args.instruction,
    resultText: result => result?.message,
    success: result => result?.success === true,
    recordSuccess: result => result?.success === true,
    recordResult: result => result,
    display: (result, resultText, args) => result?.success === true
      ? { show: true, text: `Navigated to: "${args.instruction}"`, withCheckpoint: false }
      : { show: true, text: `Navigation failed: ${resultText || 'No target selected.'}`, withCheckpoint: false }
  },
  edit_list: {
    ...COMMON_MESSAGE_MUTATION,
    executorName: 'executeEditList',
    executorArgs: args => [args.startParagraphIndex, args.endParagraphIndex, args.newItems, args.listType, args.numberingStyle],
    progressText: args => `Editing list from P${args.startParagraphIndex} to P${args.endParagraphIndex}...`,
    instruction: args => `edit_list P${args.startParagraphIndex}-P${args.endParagraphIndex}`,
  },
  insert_list_item: {
    ...COMMON_MESSAGE_MUTATION,
    executorName: 'executeInsertListItem',
    executorArgs: (args, context) => [args.afterParagraphIndex, args.text, args.indentLevel, context.docSourceBaseline],
    progressText: args => `Inserting list item after P${args.afterParagraphIndex}...`,
    instruction: args => `insert_list_item after P${args.afterParagraphIndex}`,
  },
  edit_table: {
    ...COMMON_MESSAGE_MUTATION,
    executorName: 'executeEditTable',
    executorArgs: args => [args.paragraphIndex, args.action, args.content, args.targetRow, args.targetColumn],
    progressText: args => `Editing table (${args.action})...`,
    instruction: args => `edit_table at P${args.paragraphIndex}: ${args.action}`,
  },
  edit_section: {
    ...COMMON_MESSAGE_MUTATION,
    executorName: 'executeEditSection',
    executorArgs: args => [args.sectionHeaderIndex, args.newHeaderText, args.newBodyParagraphs, args.preserveSubsections],
    progressText: args => `Editing section at P${args.sectionHeaderIndex}...`,
    instruction: args => `edit_section at P${args.sectionHeaderIndex}`,
  },
  convert_headers_to_list: {
    ...COMMON_MESSAGE_MUTATION,
    executorName: 'executeConvertHeadersToList',
    executorArgs: args => [args.paragraphIndices, args.newHeaderTexts, args.numberingFormat],
    progressText: args => `Converting ${args.paragraphIndices?.length || 0} headers to numbered list...`,
    instruction: args => `convert_headers_to_list: ${args.paragraphIndices?.join(', ')}`,
  },
  get_selection_stats: {
    executorName: 'executeGetSelectionStats', mutating: false,
    executorArgs: (args, context) => [args, context.docText ?? ''],
    progressText: () => 'Analyzing word & character count...',
    instruction: args => args.text
      ? `count text: "${args.text.substring(0, 30)}..."`
      : `count paragraphs: P${args.startParagraphIndex || 1}`,
    resultText: result => result?.message,
    success: result => !!result?.success,
    recordSuccess: result => result?.success,
    recordResult: (_result, resultText) => resultText,
    display: (result, resultText) => ({
      show: true,
      text: result?.success ? 'Calculated statistics.' : resultText,
      withCheckpoint: false
    })
  }
};

const TOOL_ENTRIES = TOOL_DEFINITIONS.map(definition => ({
  name: definition.name,
  definition,
  ...TOOL_METADATA[definition.name]
}));
const TOOL_ENTRIES_BY_NAME = new Map(TOOL_ENTRIES.map(entry => [entry.name, entry]));
const RECOVERY_TOOL_NAMES = Object.freeze(TOOL_DEFINITIONS.map(definition => definition.name));
const EMPTY_WRITE_EVIDENCE = Object.freeze({
  written: false,
  writeAttempted: false,
  status: null,
  error: null,
  receipts: null,
  operationResults: null,
  mutationOutcome: null
});

function isRecord(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function schemaTypeMatches(type, value) {
  switch (String(type || '').toUpperCase()) {
    case 'OBJECT': return isRecord(value);
    case 'ARRAY': return Array.isArray(value);
    case 'STRING': return typeof value === 'string';
    case 'INTEGER': return Number.isInteger(value);
    case 'NUMBER': return typeof value === 'number' && Number.isFinite(value);
    case 'BOOLEAN': return typeof value === 'boolean';
    default: return true;
  }
}

function validateSchemaValue(toolName, schema, value, fieldPath, errors) {
  // The table executor accepts scalar update-cell content and flat add-row
  // content, then normalizes those legacy forms with validateTableRequest.
  // Keep the public schema's broad container checks here and leave its exact
  // action-specific shape to that existing executor validator.
  const legacyTableContent = toolName === 'edit_table' && fieldPath === 'content';
  if (legacyTableContent) {
    if (typeof value !== 'string' && !Array.isArray(value)) {
      errors.push(`${fieldPath} must be a string or array.`);
    }
    return;
  }

  if (!schemaTypeMatches(schema?.type, value)) {
    errors.push(`${fieldPath} must be ${String(schema.type).toLowerCase()}.`);
    return;
  }
  if (Array.isArray(schema?.enum) && !schema.enum.includes(value)) {
    errors.push(`${fieldPath} must be one of: ${schema.enum.join(', ')}.`);
    return;
  }
  if (Array.isArray(value) && schema?.items) {
    value.forEach((item, index) => validateSchemaValue(toolName, schema.items, item, `${fieldPath}[${index}]`, errors));
  }
  if (isRecord(value) && isRecord(schema?.properties)) {
    for (const [key, propertySchema] of Object.entries(schema.properties)) {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        validateSchemaValue(toolName, propertySchema, value[key], `${fieldPath}.${key}`, errors);
      }
    }
  }
}

function validateArguments(definition, args) {
  const errors = [];
  if (!isRecord(args)) {
    return { valid: false, errors: ['Arguments must be an object.'] };
  }
  const schema = definition.parameters || {};
  if (!schemaTypeMatches(schema.type, args)) {
    return { valid: false, errors: ['Arguments must be an object.'] };
  }
  for (const field of schema.required || []) {
    if (!Object.prototype.hasOwnProperty.call(args, field)) errors.push(`${field} is required.`);
  }
  for (const [field, propertySchema] of Object.entries(schema.properties || {})) {
    if (Object.prototype.hasOwnProperty.call(args, field)) {
      validateSchemaValue(definition.name, propertySchema, args[field], field, errors);
    }
  }
  return { valid: errors.length === 0, errors };
}

function writeEvidence(result) {
  if (!isRecord(result)) return { ...EMPTY_WRITE_EVIDENCE };
  return {
    written: result.written === true,
    writeAttempted: result.writeAttempted === true,
    status: result.status ?? null,
    error: result.error ?? null,
    receipts: result.receipts ?? null,
    operationResults: result.operationResults ?? null,
    mutationOutcome: result.mutationOutcome ?? null
  };
}

function textValue(value) {
  if (typeof value === 'string') return value;
  if (value == null) return '';
  return String(value);
}

function failureResult({ name, args, metadata, code, message }) {
  const rawResult = {
    status: 'error',
    success: false,
    error: { code, message },
    message
  };
  const safeArgs = isRecord(args) ? args : {};
  let instruction;
  try { instruction = metadata?.instruction ? metadata.instruction(safeArgs) : undefined; } catch (_) { /* Keep malformed calls structured. */ }
  return {
    name,
    instruction,
    mutating: metadata?.mutating === true,
    rawResult,
    resultText: message,
    success: false,
    recordResult: message,
    recordSuccess: false,
    mutation: metadata?.mutating === true ? rawResult : null,
    hostWriteEvidence: { ...EMPTY_WRITE_EVIDENCE },
    requiresInspection: false,
    completionText: message,
    completionWithCheckpoint: false,
    display: { show: true }
  };
}

export function getToolDeclarations() {
  return JSON.parse(JSON.stringify(GEMINI_TOOLS));
}

export function createToolRegistry(executors = {}) {
  function validate(name, args) {
    const entry = TOOL_ENTRIES_BY_NAME.get(name);
    if (!entry) {
      return { valid: false, errors: [`Unknown tool: ${String(name)}.`], error: { code: 'UNKNOWN_TOOL', message: `Unknown tool: ${String(name)}.` } };
    }
    const result = validateArguments(entry.definition, args);
    return result.valid
      ? result
      : { ...result, error: { code: 'INVALID_TOOL_ARGUMENTS', message: result.errors.join(' ') } };
  }

  function describe(name, args = {}) {
    const entry = TOOL_ENTRIES_BY_NAME.get(name);
    if (!entry) return { known: false, name, progressText: 'Working...', mutating: false, withCheckpoint: false };
    const safeArgs = isRecord(args) ? args : {};
    let progressText = 'Working...';
    let instruction;
    try { progressText = entry.progressText(safeArgs); } catch (_) { /* Keep malformed calls describable. */ }
    try { instruction = entry.instruction(safeArgs); } catch (_) { /* Keep malformed calls describable. */ }
    return {
      known: true,
      name,
      mutating: entry.mutating,
      withCheckpoint: entry.mutating,
      progressText,
      instruction
    };
  }

  async function execute(name, args, context = {}) {
    const entry = TOOL_ENTRIES_BY_NAME.get(name);
    if (!entry) {
      return failureResult({
        name,
        args,
        code: 'UNKNOWN_TOOL',
        message: `Unknown tool: ${String(name)}.`
      });
    }

    const validation = validate(name, args);
    if (!validation.valid) {
      return failureResult({
        name,
        args,
        metadata: entry,
        code: validation.error.code,
        message: validation.error.message
      });
    }

    const implementation = executors?.[entry.executorName];
    if (typeof implementation !== 'function') {
      const message = `Tool executor unavailable: ${entry.executorName}.`;
      return failureResult({ name, args, metadata: entry, code: 'TOOL_EXECUTOR_UNAVAILABLE', message });
    }

    if (typeof entry.executorArgs !== 'function') {
      return failureResult({ name, args, metadata: entry, code: 'TOOL_EXECUTOR_MAPPING_MISSING', message: `No executor argument mapping exists for ${name}.` });
    }
    const rawResult = await implementation(...entry.executorArgs(args, context));

    const resultText = textValue(entry.resultText(rawResult));
    const success = !!entry.success(rawResult);
    const recordSuccess = entry.recordSuccess(rawResult);
    const display = entry.display(rawResult, resultText, args);
    const evidence = entry.mutating ? writeEvidence(rawResult) : { ...EMPTY_WRITE_EVIDENCE };
    const requiresInspection = entry.mutating
      && ['indeterminate', 'partial', 'applied_with_host_error'].includes(evidence.mutationOutcome);

    return {
      name,
      instruction: entry.instruction(args),
      mutating: entry.mutating,
      rawResult,
      resultText,
      success,
      recordResult: entry.recordResult(rawResult, resultText),
      recordSuccess,
      mutation: entry.mutating ? rawResult : null,
      hostWriteEvidence: evidence,
      requiresInspection,
      completionText: display.text,
      completionWithCheckpoint: display.withCheckpoint === true,
      display: { show: display.show === true }
    };
  }

  return {
    recoveryToolNames: RECOVERY_TOOL_NAMES,
    validate,
    describe,
    execute
  };
}
