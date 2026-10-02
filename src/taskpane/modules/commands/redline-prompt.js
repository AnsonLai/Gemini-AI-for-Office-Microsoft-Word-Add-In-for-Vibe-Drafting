/**
 * redline-prompt.js
 *
 * Shared, pure (no Office.js / DOM) builder for the redline diff-generation
 * prompt and its response schema. Both the Word add-in (agentic-tools.js) and
 * the offline eval harness (tests/evals) import these so they exercise the exact
 * same prompt. Do NOT change the prompt text here without intending to change
 * model behavior everywhere.
 */

/**
 * Build the diff-generation prompt for a redline request.
 * @param {string} instruction - the user edit instruction
 * @param {string} fullDocumentText - the document content with [P#] anchors
 * @returns {string}
 */
export function buildRedlineDiffPrompt(instruction, fullDocumentText) {
  return `You are an expert legal editor. Review the document content (provided with [P#] anchors) based on the user's instruction.
Generate a JSON array of precise changes to be made, referencing the paragraph numbers.

CRITICAL: Return ONLY valid JSON. Do NOT include explanatory text, notes, or duplicate entries.

Each change must be an object with the following structure:
- "paragraphIndex": The integer number of the paragraph to modify (e.g., 1 for [P1]). For "replace_range", this is the START paragraph.
- "anchorText": REQUIRED for every change. Copy the first 30-60 characters of the CURRENT text of the paragraph at "paragraphIndex" VERBATIM from the document content above (do NOT include the [P#] marker). For "replace_range", anchor the START paragraph. For an empty paragraph, use "". This is used to verify your edit lands on the correct paragraph.
- "endParagraphIndex": (Only for "replace_range") The integer number of the END paragraph (inclusive).
- "operation": "edit_paragraph", "replace_paragraph", "modify_text", "replace_range", or "format_text".
- "newContent": (For "edit_paragraph" full rewrites ONLY) The complete rewritten paragraph content.
- "replacements": (For "edit_paragraph" localized edits ONLY) An array of {"find":"exact existing text","replace":"new text"}. Use this instead of "newContent" for word, phrase, or sentence edits within one paragraph. Each "find" must match the CURRENT paragraph text exactly, including case and punctuation. Use "occurrence" (1-based) if "find" appears more than once.
- "content": (REQUIRED for "replace_paragraph" and "replace_range" ONLY) The new content to insert.
- "originalText": (For "modify_text" ONLY) The specific text snippet within the paragraph to find and replace. **MAX 80 characters**.
- "replacementText": (For "modify_text" ONLY) The new text to replace "originalText" with.
- "find": (For "format_text" ONLY) Exact single-line text in the paragraph whose character formatting changes. If it appears more than once, add the 1-based "occurrence" of the appearance to format.
- "formatting": (For "format_text" ONLY) An object of true/false flags among "bold", "italic", "underline", "strikethrough". false REMOVES that formatting, e.g. {"bold":false}.

**MARKDOWN FORMATTING (VERY IMPORTANT)**:
All content, newContent, replacementText, and replacements.replace values support Markdown formatting. Use these when the user requests formatting:
- **Bold**: Use **text** (double asterisks)
- *Italic*: Use *text* (single asterisks)
- **Underline**: Use ++text++ (double pluses)
- ~~Strikethrough~~: Use ~~text~~ (double tildes)
- ***Bold Italic***: Use ***text*** (triple asterisks)
- **Unordered/Bullet lists**: Use "- item" or "* item" on separate lines. These render as bullet points (•).
- **Ordered/Numbered lists**: Use "1. item", "2. item" on separate lines. These render as 1, 2, 3...
- **Alphabetical lists (A, B, C)**: Use "A. item", "B. item" on separate lines. Use lowercase "a. item" for a, b, c. Use "I.", "II." for roman numerals.
- Line breaks: Use actual newlines (\\n) in the text
- Tables: Use GitHub-style markdown tables:
  | Header 1 | Header 2 |
  |----------|----------|
  | Cell 1   | Cell 2   |
- Headings: Use # for H1, ## for H2, ### for H3

**CRITICAL LIST FORMATTING RULES**:
- **PRESERVE HIERARCHY**: If the document uses nested numbering (1.1, 1.1.1, etc.), ALWAYS use that same hierarchical format in your changes. **Do NOT flatten nested lists** into simple numbered lists (1., 2., 3.) unless specifically asked to restructure the hierarchy.
- **INCLUDE MARKERS**: Always include the correct list marker (e.g., "1.1.1 ") at the start of your \`newContent\` or \`content\` for list items. The system will use these to correctly set the indentation level in Word, and then it will automatically strip them from the final text.
- **NO MIXING**: NEVER mix bullet markers with manual numbering like "• (a)" or "- 1." - this creates malformed output
- **MARKDOWN SYNTAX**: 
  - For bullets: use "- " or "* "
  - For simple numbers: use "1. ", "2. "
  - For hierarchical numbers: use "1.1. ", "1.1.1. "
- **STRIPPING**: When converting existing lists, REMOVE the original markers from your response and use ONLY the markdown syntax described above.

When the user asks for formatted content (bullets, tables, bold, etc.), ALWAYS use the appropriate Markdown syntax.

**FORMAT-ONLY CHANGES (unbold, un-italicize, remove underline, or format existing text without changing it)**:
- Use "format_text" and never rewrite the text: unchanged text without Markdown markers does NOT remove existing formatting.
- Example removing bold: {"paragraphIndex":12,"anchorText":"The Receiving Party shall","operation":"format_text","find":"Confidential Information","formatting":{"bold":false}}
- Use one "format_text" change per distinct span; several may target the same paragraph.

Rules:
- **PREFER LOCALIZED EDITS**: For word, phrase, or sentence edits inside one existing paragraph, use "edit_paragraph" with "replacements": [{"find":"exact current text","replace":"new text"}] instead of rewriting the entire paragraph. Multiple distinct replacements may be included in one change.
- Use "edit_paragraph" with "newContent" for broad paragraph rewrites. Provide either "replacements" or "newContent", never both.
- Use "replace_paragraph" only when you need to replace with complex formatted content (lists, tables, headings) that requires HTML insertion.
- If converting text into a Markdown table:
  - Use "replace_paragraph" when it is a single paragraph.
  - Use "replace_range" when it spans multiple consecutive paragraphs.
  - For consecutive source paragraphs, ALWAYS include "endParagraphIndex" covering the entire source block being replaced.
  - Do NOT use "modify_text" for table conversions.
  - The "content" value MUST be a complete multiline Markdown table with a header row, separator row, and data rows.
  - NEVER return a single pipe-delimited line such as "A|B|C"; that is plain text, not a table.
  - Example valid table content: "| Column A | Column B |\\n|---|---|\\n| Value A | Value B |"
  - For multi-line source blocks, use additional table rows instead of HTML tags inside cells.
  - Example for turning party paragraphs into a two-column table:
    [{"paragraphIndex":4,"endParagraphIndex":6,"operation":"replace_range","content":"| Disclosing Party | Receiving Party |\\n|---|---|\\n| [Name of Disclosing Party] | [Name of Receiving Party] |\\n| [Address of Disclosing Party] | [Address of Receiving Party] |"}]
- Use legacy "modify_text" only when the localized "edit_paragraph" form is unavailable.
- Never use "modify_text" when the replacement includes line breaks, list markers, headings, or Markdown tables.
- **CRITICAL LENGTH LIMIT**: For "modify_text", "originalText" MUST be **80 characters or fewer**. This is a hard limit.
- Use "replace_range" when you need to replace multiple consecutive paragraphs (like converting a bulleted list to a single paragraph).
- For "replace_range", provide ONLY "paragraphIndex", "endParagraphIndex", "operation", and "content". Do NOT include "originalText" or "replacementText".
- A "replace_range" or "replace_paragraph" without a non-empty "content" field is INVALID. If you cannot determine the replacement content, return [].
- INVALID replace_range example: {"paragraphIndex":3,"operation":"replace_range","endParagraphIndex":5,"originalText":"","replacementText":""}
- Never put schema explanations, validation errors, or instructions about JSON fields inside "content", "newContent", or "replacementText". These fields must contain ONLY text that should appear in the Word document.
- For "edit_paragraph", provide "paragraphIndex", "operation", "anchorText", and exactly one of "replacements" or "newContent".
- For localized replacements, copy each "find" verbatim from the target paragraph. If the text occurs more than once, specify the 1-based "occurrence". Do not use localized replacements for multi-paragraph edits or structural content such as lists and tables.
- For "modify_text", "originalText" must match EXACTLY text found within that specific paragraph.
- Do NOT include the [P#] marker in any content fields.
- Return ONLY ONE change per unique text location. Do NOT create duplicate entries.
- NEVER write notes, schema commentary, or your own reasoning inside ANY field value. Every string field must contain ONLY document text (or be omitted). OMIT fields that do not apply to the chosen operation.
- **ADDING NEW CONTENT**: Use ONE change with ALL the new paragraphs in its "content", separated by newlines. Do NOT spread new paragraphs across multiple paragraph indices.
  - If there is a blank/empty paragraph where the content belongs, target it: {"paragraphIndex":5,"operation":"replace_paragraph","anchorText":"","content":"Adaptive solutions\\nOur solutions evolve with your project.\\nDedicated partnership"}
  - To add content at the very END of the document, target the paragraph number ONE PAST the last [P#] (e.g. if the document ends at [P5], use paragraphIndex 6) with "replace_paragraph", "anchorText":"", and all new paragraphs in "content". This appends them after the last paragraph.
  - Otherwise NEVER use a paragraphIndex more than one beyond the last [P#] shown.

IMPORTANT: This document may contain existing tracked changes. The text shown represents the "accepted" state (as if all changes were accepted). Your changes will be applied as additional tracked changes on top of existing ones.

USER INSTRUCTION:
"${instruction}"

DOCUMENT CONTENT:
"""${fullDocumentText}"""

Return ONLY the JSON array, nothing else:`;
}

// Gemini responseSchema for the diff-generation call.
export const REDLINE_DIFF_SCHEMA = {
  type: "ARRAY",
  items: {
    type: "OBJECT",
    properties: {
      paragraphIndex: { type: "INTEGER", description: "The paragraph number (1-based)" },
      anchorText: {
        type: "STRING",
        description:
          'First 30-60 characters of the CURRENT text of the targeted paragraph, copied verbatim from the document content (without the [P#] marker). Used to verify the edit targets the correct paragraph. Use "" for an empty paragraph.',
      },
      endParagraphIndex: {
        type: "INTEGER",
        description:
          "Only for replace_range: the end paragraph number (inclusive). Required when converting multiple paragraphs into a table.",
      },
      operation: {
        type: "STRING",
        enum: ["edit_paragraph", "replace_paragraph", "modify_text", "replace_range", "format_text"],
        description: "The type of operation to perform",
      },
      newContent: {
        type: "STRING",
        description: "For edit_paragraph full rewrites only: the complete rewritten paragraph content. Omit when using replacements.",
      },
      replacements: {
        type: "ARRAY",
        description: "For localized edit_paragraph changes inside one paragraph. Use instead of newContent. Each find must exactly match current paragraph text.",
        items: {
          type: "OBJECT",
          properties: {
            find: { type: "STRING", description: "Exact, case-sensitive text in the current paragraph" },
            replace: { type: "STRING", description: "Replacement text; may be empty to delete the found text" },
            occurrence: { type: "INTEGER", description: "1-based occurrence when find appears more than once" },
          },
          required: ["find", "replace"],
        },
      },
      find: {
        type: "STRING",
        description: "ONLY for format_text: exact single-line text in the paragraph whose character formatting changes. Must occur exactly once unless occurrence is given.",
      },
      occurrence: {
        type: "INTEGER",
        description: "ONLY for format_text: 1-based occurrence of a repeated find to format.",
      },
      formatting: {
        type: "OBJECT",
        description: "ONLY for format_text: character-format flags to set; false removes the formatting (e.g. bold false to unbold).",
        properties: {
          bold: { type: "BOOLEAN" },
          italic: { type: "BOOLEAN" },
          underline: { type: "BOOLEAN" },
          strikethrough: { type: "BOOLEAN" },
        },
      },
      content: {
        type: "STRING",
        description:
          "Required for replace_paragraph and replace_range: the new content. For tables, use a complete multiline GitHub Markdown table with header, separator, and data rows. Do not use replacementText for these operations.",
      },
      originalText: {
        type: "STRING",
        description:
          "ONLY for modify_text: the text to find (max 80 chars). OMIT this field entirely for every other operation. Never put notes or reasoning here.",
      },
      replacementText: {
        type: "STRING",
        description:
          "ONLY for modify_text: the replacement text. OMIT this field entirely for every other operation. Never put notes, explanations, or reasoning here.",
      },
    },
    required: ["paragraphIndex", "operation", "anchorText"],
  },
};

/**
 * Build a one-shot corrective retry prompt after a change set was rejected.
 * Shows the model its own invalid output plus the machine-generated rejection
 * reasons, and asks for a corrected JSON array against the same document.
 *
 * @param {string} basePrompt - the original prompt from buildRedlineDiffPrompt
 * @param {Array} previousChanges - the raw (invalid) change array the model returned
 * @param {string} rejectionDetail - formatRejections() output for those changes
 * @returns {string}
 */
export function buildCorrectiveRetryPrompt(basePrompt, previousChanges, rejectionDetail) {
  let previousJson;
  try {
    previousJson = JSON.stringify(previousChanges, null, 2);
  } catch {
    previousJson = String(previousChanges);
  }
  // A degenerate previous response (e.g. one change repeated dozens of times)
  // must not blow up the retry prompt.
  const MAX_PREVIOUS_JSON_CHARS = 4000;
  if (previousJson && previousJson.length > MAX_PREVIOUS_JSON_CHARS) {
    previousJson = `${previousJson.slice(0, MAX_PREVIOUS_JSON_CHARS)}\n… (truncated; the response continued repeating similar entries)`;
  }
  return `${basePrompt}

YOUR PREVIOUS RESPONSE WAS REJECTED. You returned:
${previousJson}

It failed validation for these reasons:
${rejectionDetail}

Produce a CORRECTED JSON array for the same instruction and document above. Rules for the correction:
- Fix ONLY the problems listed; keep the intent of the edit.
- Include every REQUIRED field for the chosen operation ("content" for replace_paragraph/replace_range; exactly one of "newContent" or non-empty "replacements" for edit_paragraph).
- Content fields must contain ONLY the text that should appear in the document. NEVER write notes, schema commentary, or reasoning inside any field value.
- OMIT fields that do not apply to the chosen operation.

Return ONLY the corrected JSON array, nothing else:`;
}
