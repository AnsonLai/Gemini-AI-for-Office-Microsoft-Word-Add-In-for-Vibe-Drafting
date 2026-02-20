/* global Word */

import {
  applyRedlineToOxml,
  preprocessMarkdown,
  ReconciliationPipeline,
  wrapInDocumentFragment,
  getAuthorForTracking,
  buildListMarkdown,
  normalizeListItemsWithLevels,
  withNativeTrackingDisabled,
  applySharedOperationToWordParagraph,
  applySharedOperationToWordScope,
  applyRedlineChangesToWordContext
} from '../reconciliation/index.js';
import {
  detectDocumentFont
} from '../utils/markdown-utils.js';

let loadApiKey;
let loadModel;
let loadSystemMessage;
let loadRedlineSetting;
let loadRedlineAuthor;
let setChangeTrackingForAi;
let restoreChangeTracking;
let SAFETY_SETTINGS_BLOCK_NONE;
let API_LIMITS;

function initAgenticTools(deps) {
  ({
    loadApiKey,
    loadModel,
    loadSystemMessage,
    loadRedlineSetting,
    loadRedlineAuthor,
    setChangeTrackingForAi,
    restoreChangeTracking,
    SAFETY_SETTINGS_BLOCK_NONE,
    API_LIMITS
  } = deps);
}

/**
 * Agentic Tool: Applies redlines based on an instruction using Structural Anchoring.
 */
async function executeRedline(instruction, fullDocumentText) {
  // Check for API key
  const geminiApiKey = loadApiKey();
  if (!geminiApiKey) {
    return "Error: Please set your Gemini API key in the Settings.";
  }

  try {
    // Detect document font for consistent HTML insertion
    await detectDocumentFont();
    // 1. Build the prompt for the diff generator
    const fullPrompt = `You are an expert legal editor. Review the document content (provided with [P#] anchors) based on the user's instruction.
Generate a JSON array of precise changes to be made, referencing the paragraph numbers.

CRITICAL: Return ONLY valid JSON. Do NOT include explanatory text, notes, or duplicate entries.

Each change must be an object with the following structure:
- "paragraphIndex": The integer number of the paragraph to modify (e.g., 1 for [P1]). For "replace_range", this is the START paragraph.
- "endParagraphIndex": (Only for "replace_range") The integer number of the END paragraph (inclusive).
- "operation": "edit_paragraph", "replace_paragraph", "modify_text", or "replace_range".
- "newContent": (For "edit_paragraph" ONLY) The complete rewritten paragraph content. The system will automatically compute precise word-level changes.
- "content": (For "replace_paragraph" and "replace_range" ONLY) The new content to insert.
- "originalText": (For "modify_text" ONLY) The specific text snippet within the paragraph to find and replace. **MAX 80 characters**.
- "replacementText": (For "modify_text" ONLY) The new text to replace "originalText" with.

**MARKDOWN FORMATTING (VERY IMPORTANT)**:
All content and replacementText values support Markdown formatting. Use these when the user requests formatting:
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

Rules:
- **PRIORITIZE \`edit_paragraph\`**: This is the NEW preferred method. For ANY text edit (small or large), use \`edit_paragraph\` with the complete rewritten paragraph. The system will automatically compute precise word-level changes using diff-match-patch. This is more reliable than \`modify_text\`.
- Use "edit_paragraph" for ALL text edits: spelling changes, word replacements, sentence rewrites, or even 60% paragraph rewrites. Just provide the full new paragraph content.
- Use "replace_paragraph" only when you need to replace with complex formatted content (lists, tables, headings) that requires HTML insertion.
- If converting text into a Markdown table:
  - Use "replace_paragraph" when it is a single paragraph.
  - Use "replace_range" when it spans multiple consecutive paragraphs.
  - Do NOT use "modify_text" for table conversions.
- Use "modify_text" ONLY as a fallback for very specific surgical edits where you need to target exact substrings.
- Never use "modify_text" when the replacement includes line breaks, list markers, headings, or Markdown tables.
- **CRITICAL LENGTH LIMIT**: For "modify_text", "originalText" MUST be **80 characters or fewer**. This is a hard limit.
- Use "replace_range" when you need to replace multiple consecutive paragraphs (like converting a bulleted list to a single paragraph).
- For "replace_range", provide ONLY "paragraphIndex", "endParagraphIndex", "operation", and "content". Do NOT include "originalText" or "replacementText".
- For "edit_paragraph", provide ONLY "paragraphIndex", "operation", and "newContent".
- For "modify_text", "originalText" must match EXACTLY text found within that specific paragraph.
- Do NOT include the [P#] marker in any content fields.
- Return ONLY ONE change per unique text location. Do NOT create duplicate entries.

IMPORTANT: This document may contain existing tracked changes. The text shown represents the "accepted" state (as if all changes were accepted). Your changes will be applied as additional tracked changes on top of existing ones.

USER INSTRUCTION:
"${instruction}"

DOCUMENT CONTENT:
"""${fullDocumentText}"""

Return ONLY the JSON array, nothing else:`;

    // 2. Call Gemini to get the JSON array of changes
    const aiChanges = await callGeminiForDiffs(fullPrompt);

    console.log("AI Suggested Changes (raw):", aiChanges);

    if (!aiChanges || !Array.isArray(aiChanges)) {
      return {
        message: "AI did not return a valid list of changes. Please check the console logs for details.",
        showToUser: false  // Silent error - let the model handle it
      };
    }

    if (aiChanges.length === 0) {
      return {
        message: "AI had no changes to suggest based on the instruction.",
        showToUser: false  // Silent - let the model try again or respond
      };
    }

    {
      const redlineEnabled = loadRedlineSetting();
      const redlineAuthor = loadRedlineAuthor();
      let changesApplied = 0;

      await Word.run(async (context) => {
        const trackingState = await setChangeTrackingForAi(context, redlineEnabled, "executeRedline");
        try {
          context.document.load("changeTrackingMode");
          await context.sync();
          const baseTrackingMode = context.document.changeTrackingMode;
          const result = await applyRedlineChangesToWordContext(context, aiChanges, {
            author: redlineAuthor,
            generateRedlines: redlineEnabled,
            disableNativeTracking: redlineEnabled,
            baseTrackingMode,
            logPrefix: "Redline/Shared"
          });
          changesApplied = result.changesApplied;
        } finally {
          await restoreChangeTracking(context, trackingState, "executeRedline");
        }
      });

      if (changesApplied === 0) {
        return {
          message: "Applied 0 edits. The AI's suggestions could not be mapped to the document content.",
          showToUser: false
        };
      }

      return {
        message: `Successfully applied ${changesApplied} edits${redlineEnabled ? ' with redlines' : ' without redlines'}.`,
        showToUser: true
      };
    }

  } catch (error) {
    console.error("Error in executeRedline:", error);
    return {
      message: `Error applying redlines: ${error.message}`,
      showToUser: false  // Silent error - let the model handle it
    };
  }
}
// Helper for the Diff generation (specialized prompt)
async function callGeminiForDiffs(prompt) {
  const geminiApiKey = loadApiKey();
  const geminiModel = loadModel();
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${geminiModel}:generateContent?key=${geminiApiKey}`;

  const jsonSchema = {
    type: "ARRAY",
    items: {
      type: "OBJECT",
      properties: {
        "paragraphIndex": { "type": "INTEGER", "description": "The paragraph number (1-based)" },
        "endParagraphIndex": { "type": "INTEGER", "description": "Only for replace_range: the end paragraph number (inclusive)" },
        "operation": {
          "type": "STRING",
          "enum": ["edit_paragraph", "replace_paragraph", "modify_text", "replace_range"],
          "description": "The type of operation to perform"
        },
        "newContent": { "type": "STRING", "description": "For edit_paragraph only: the complete rewritten paragraph content" },
        "content": { "type": "STRING", "description": "For replace_paragraph and replace_range: the new content" },
        "originalText": { "type": "STRING", "description": "For modify_text only: the text to find (max 80 chars). Split larger edits into multiple operations." },
        "replacementText": { "type": "STRING", "description": "For modify_text only: the replacement text" }
      },
      required: ["paragraphIndex", "operation"]
    }
  };

  const systemInstruction = {
    parts: [
      {
        text: loadSystemMessage(),
      },
    ],
  };

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    systemInstruction: systemInstruction,
    safetySettings: SAFETY_SETTINGS_BLOCK_NONE,
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: API_LIMITS.MAX_OUTPUT_TOKENS,
      responseMimeType: "application/json",
      responseSchema: jsonSchema,
    },
  };

  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`API failed: ${err}`);
    }

    const result = await response.json();
    console.log("Gemini diff raw result:", JSON.stringify(result, null, 2));

    if (!result.candidates || !Array.isArray(result.candidates) || result.candidates.length === 0) {
      throw new Error("Gemini diff response contained no candidates.");
    }

    const candidate = result.candidates[0];

    if (!candidate.content || !candidate.content.parts || !Array.isArray(candidate.content.parts) || candidate.content.parts.length === 0) {
      console.error("Gemini diff candidate missing content.parts:", candidate);
      throw new Error("Gemini diff response was missing content.parts (possibly blocked by safety settings).");
    }

    const jsonText = candidate.content.parts[0].text;
    console.log("Gemini diff JSON text:", jsonText);
    return JSON.parse(jsonText);
  } catch (error) {
    console.error("Error getting diffs:", error);
    return null;
  }
}
/**
 * Agentic Tool: Inserts comments based on an instruction using Structural Anchoring.
 */
async function executeComment(instruction, fullDocumentText) {
  const geminiApiKey = loadApiKey();
  if (!geminiApiKey) {
    return "Error: Please set your Gemini API key in the Settings.";
  }

  try {
    const fullPrompt = `You are an expert legal editor. Review the document content (provided with [P#] anchors) based on the user's instruction.
Generate a JSON array of comments to be inserted, referencing the paragraph numbers.

Each item must be an object with:
- "paragraphIndex": The integer number of the paragraph to comment on (e.g., 1 for [P1]).
- "textToFind": The specific text snippet within the paragraph to attach the comment to. Must match EXACTLY. CRITICAL: Keep this VERY SHORT - maximum 50 characters or 5-8 words. Use a unique phrase that identifies the location.
- "commentContent": The text of the comment.

USER INSTRUCTION:
"${instruction}"

DOCUMENT CONTENT:
"""${fullDocumentText}"""

JSON ARRAY OF COMMENTS:`;

    const aiComments = await callGeminiForJSON(fullPrompt, {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          "paragraphIndex": { "type": "INTEGER" },
          "textToFind": { "type": "STRING" },
          "commentContent": { "type": "STRING" }
        },
        required: ["paragraphIndex", "textToFind", "commentContent"]
      }
    });
    console.log("AI Suggested Comments:", aiComments);

    if (!aiComments || !Array.isArray(aiComments) || aiComments.length === 0) {
      return {
        message: "AI had no comments to suggest.",
        showToUser: false  // Silent - let the model try again or respond
      };
    }

    let commentsApplied = 0;

    await Word.run(async (context) => {
      const redlineEnabled = loadRedlineSetting();
      const redlineAuthor = loadRedlineAuthor();
      const trackingState = await setChangeTrackingForAi(context, false, "executeComment");

      try {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items/text, items/style");
        await context.sync();

        for (const item of aiComments) {
          const pIndex = item.paragraphIndex - 1;
          if (pIndex < 0 || pIndex >= paragraphs.items.length) continue;

          const targetParagraph = paragraphs.items[pIndex];
          try {
            const applied = await applySharedOperationToWordParagraph({
              context,
              targetParagraph,
              operation: {
                type: "comment",
                targetRef: "P1",
                target: targetParagraph.text || item.textToFind || "",
                textToComment: item.textToFind,
                commentContent: item.commentContent
              },
              author: redlineAuthor,
              generateRedlines: redlineEnabled,
              logPrefix: "Comment/Shared"
            });

            if (applied) {
              commentsApplied += 1;
              console.log(`[Comment/Shared] Applied comment via shared engine in P${item.paragraphIndex}`);
            } else {
              console.warn(`[Comment/Shared] No changes produced for P${item.paragraphIndex}`);
            }
          } catch (sharedError) {
            console.warn(`[Comment/Shared] Failed in P${item.paragraphIndex} (no fallback):`, sharedError?.message || sharedError);
          }
        }
      } finally {
        await restoreChangeTracking(context, trackingState, "executeComment");
      }
    });

    return createToolResult(commentsApplied, 'comments', "Inserted 0 comments. The AI's suggestions could not be mapped to the document content.");

  } catch (error) {
    console.error("Error in executeComment:", error);
    return {
      message: `Error inserting comments: ${error.message}`,
      showToUser: false  // Silent error - let the model handle it
    };
  }
}
/**
 * Agentic Tool: Highlights text based on an instruction using Structural Anchoring.
 * @param {string} instruction - The instruction for what to highlight
 * @param {string} fullDocumentText - The document content with paragraph anchors
 * @param {string} highlightColor - The default highlight color (default: "Yellow")
 */
async function executeHighlight(instruction, fullDocumentText, highlightColor = "Yellow") {
  const geminiApiKey = loadApiKey();
  if (!geminiApiKey) {
    return "Error: Please set your Gemini API key in the Settings.";
  }

  // Normalize color to proper case for Word API
  const normalizedColor = highlightColor.charAt(0).toUpperCase() + highlightColor.slice(1).toLowerCase();

  try {
    const fullPrompt = `You are an expert legal editor. Review the document content (provided with [P#] anchors) based on the user's instruction.
Generate a JSON array of highlights to be applied, referencing the paragraph numbers.

Each item must be an object with:
- "paragraphIndex": The integer number of the paragraph (e.g., 1 for [P1]).
- "textToFind": The specific text snippet within the paragraph to highlight. Must match EXACTLY. CRITICAL: Keep this VERY SHORT - maximum 50 characters or 5-8 words. Use a unique phrase that identifies the location.

USER INSTRUCTION:
"${instruction}"

DOCUMENT CONTENT:
"""${fullDocumentText}"""

JSON ARRAY OF HIGHLIGHTS:`;

    const aiHighlights = await callGeminiForJSON(fullPrompt, {
      type: "ARRAY",
      items: {
        type: "OBJECT",
        properties: {
          "paragraphIndex": { "type": "INTEGER" },
          "textToFind": { "type": "STRING" }
        },
        required: ["paragraphIndex", "textToFind"]
      }
    });
    console.log("AI Suggested Highlights:", aiHighlights);

    if (!aiHighlights || !Array.isArray(aiHighlights) || aiHighlights.length === 0) {
      return {
        message: "AI had no highlights to suggest.",
        showToUser: false  // Silent - let the model try again or respond
      };
    }

    let highlightsApplied = 0;

    await Word.run(async (context) => {
      // Load redline settings
      const redlineEnabled = loadRedlineSetting();
      const authorName = getAuthorForTracking();

      // CRITICAL: For OOXML insertion with manual redline tags (w:rPrChange), 
      // we MUST turn OFF Word's native Track Changes during the insertion.
      // Otherwise Word will redline the entire paragraph replacement.
      const trackingState = await setChangeTrackingForAi(context, false, "executeHighlight");

      try {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items/text, items/style");
        await context.sync();

        for (const item of aiHighlights) {
          const pIndex = item.paragraphIndex - 1;
          if (pIndex < 0 || pIndex >= paragraphs.items.length) continue;

          const targetParagraph = paragraphs.items[pIndex];
          try {
            const applied = await applySharedOperationToWordParagraph({
              context,
              targetParagraph,
              operation: {
                type: "highlight",
                targetRef: "P1",
                target: targetParagraph.text || item.textToFind || "",
                textToHighlight: item.textToFind,
                color: normalizedColor.toLowerCase()
              },
              author: authorName,
              generateRedlines: redlineEnabled,
              logPrefix: "Highlight/Shared"
            });

            if (applied) {
              highlightsApplied++;
              console.log(`[Highlight/Shared] Applied ${normalizedColor} highlight to "${item.textToFind}" in P${item.paragraphIndex}`);
            } else {
              console.warn(`[Highlight/Shared] No changes produced for P${item.paragraphIndex}`);
            }
          } catch (sharedError) {
            console.warn(`[Highlight/Shared] Failed in P${item.paragraphIndex} (no fallback):`, sharedError?.message || sharedError);
          }
        }
      } finally {
        await restoreChangeTracking(context, trackingState, "executeHighlight");
      }
    });

    return createToolResult(highlightsApplied, 'highlights', "Highlighted 0 items. The AI's suggestions could not be mapped to the document content.");

  } catch (error) {
    console.error("Error in executeHighlight:", error);
    return {
      message: `Error highlighting text: ${error.message}`,
      showToUser: false
    };
  }
}
/**
 * Agentic Tool: Navigates to and selects a specific section of the document.
 */
async function executeNavigate(instruction, fullDocumentText) {
  const geminiApiKey = loadApiKey();
  if (!geminiApiKey) {
    return "Error: Please set your Gemini API key in the Settings.";
  }

  try {
    const fullPrompt = `You are an expert document navigator. Review the document content (provided with [P#] anchors) based on the user's navigation instruction.
Determine the most relevant paragraph to navigate to and provide navigation details.

Return a JSON object with:
- "paragraphIndex": The integer number of the paragraph to navigate to (e.g., 1 for [P1]).
- "navigationDescription": A brief description of what was found and where the user was taken (e.g., "Navigated to paragraph 3: Introduction section", "Found the signature block at paragraph 15").

USER INSTRUCTION:
"${instruction}"

DOCUMENT CONTENT:
"""${fullDocumentText}"""

JSON RESPONSE:`;

    const navigationResult = await callGeminiForJSON(fullPrompt, {
      type: "OBJECT",
      properties: {
        "paragraphIndex": { "type": "INTEGER" },
        "navigationDescription": { "type": "STRING" }
      },
      required: ["paragraphIndex"]
    });
    console.log("AI Navigation Result:", navigationResult);

    if (!navigationResult || !navigationResult.paragraphIndex) {
      return {
        message: "Could not determine where to navigate based on the instruction.",
        showToUser: false
      };
    }

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();

      const pIndex = navigationResult.paragraphIndex - 1;
      if (pIndex < 0 || pIndex >= paragraphs.items.length) {
        throw new Error(`Invalid paragraph index: ${navigationResult.paragraphIndex}`);
      }

      const targetParagraph = paragraphs.items[pIndex];

      // Select the paragraph to navigate to it
      targetParagraph.select();
      await context.sync();
    });

    const description = navigationResult.navigationDescription || `Navigated to paragraph ${navigationResult.paragraphIndex}`;

    return {
      message: description,
      showToUser: true
    };

  } catch (error) {
    console.error("Error in executeNavigate:", error);
    return {
      message: `Error navigating: ${error.message}`,
      showToUser: false
    };
  }
}
// ==================== TOOL EXECUTION HELPERS ====================

/**
 * Validates that prerequisites for tool execution are met (API key exists).
 * @returns {Object} Object with either { apiKey } or { error }
 */
function validateToolPrerequisites() {
  const apiKey = loadApiKey();
  if (!apiKey) {
    return { error: "Error: Please set your Gemini API key in the Settings." };
  }
  return { apiKey };
}

/**
 * Creates a standardized tool execution result object.
 * @param {number} count - Number of items successfully processed
 * @param {string} itemType - Type of item (e.g., "comments", "highlights")
 * @param {string} zeroMessage - Optional custom message for zero count
 * @returns {Object} Result object with { message, showToUser }
 */
function createToolResult(count, itemType, zeroMessage) {
  if (count === 0) {
    return {
      message: zeroMessage || `Applied 0 ${itemType}. The AI's suggestions could not be mapped to the document content.`,
      showToUser: false  // Silent fallback
    };
  }

  const actionVerb = itemType === 'comments' ? 'inserted' : itemType === 'highlights' ? 'highlighted' : 'applied';
  return {
    message: `Successfully ${actionVerb} ${count} ${itemType}.`,
    showToUser: true
  };
}

// Generic helper for JSON responses
async function callGeminiForJSON(prompt, schema) {
  const geminiApiKey = loadApiKey();
  const geminiModel = loadModel();
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${geminiModel}:generateContent?key=${geminiApiKey}`;

  const systemInstruction = {
    parts: [{ text: loadSystemMessage() }]
  };

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    systemInstruction: systemInstruction,
    safetySettings: SAFETY_SETTINGS_BLOCK_NONE,
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: 48000,
      responseMimeType: "application/json",
      responseSchema: schema,
    },
  };

  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`API failed: ${err}`);
    }

    const result = await response.json();
    if (!result.candidates || result.candidates.length === 0) throw new Error("No candidates");
    const candidate = result.candidates[0];
    if (!candidate.content || !candidate.content.parts) throw new Error("No content");

    const jsonText = candidate.content.parts[0].text;
    return JSON.parse(jsonText);
  } catch (error) {
    console.error("Error calling Gemini for JSON:", error);
    return null;
  }
}


async function executeResearch(query) {
  const geminiApiKey = loadApiKey();
  const geminiModel = loadModel();
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${geminiModel}:generateContent?key=${geminiApiKey}`;

  const tools = [{ google_search: {} }];

  const payload = {
    contents: [{ parts: [{ text: query }] }],
    tools: tools,
    safetySettings: [
      { category: "HARM_CATEGORY_HARASSMENT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_HATE_SPEECH", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold: "BLOCK_NONE" },
      { category: "HARM_CATEGORY_DANGEROUS_CONTENT", threshold: "BLOCK_NONE" }
    ]
  };

  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`Research API failed: ${err}`);
    }

    const result = await response.json();
    if (!result.candidates || result.candidates.length === 0) return "No results found.";

    const candidate = result.candidates[0];
    if (!candidate.content || !candidate.content.parts) return "No content returned.";

    return candidate.content.parts[0].text;
  } catch (error) {
    console.error("Error in executeResearch:", error);
    return `Error performing research: ${error.message}`;
  }
}

/**
 * Maintains a rolling window of chat history while preserving function call/response pairs
 */

/**
 * Execute insert_list_item tool - surgically insert a single list item after a specific paragraph
 * @param {number} afterParagraphIndex - 1-based paragraph index to insert after
 * @param {string} text - The text content (without numbering)
 * @param {number} indentLevel - Relative indent: 0=same, 1=deeper, -1=shallower
 */
/**
 * Execute insert_list_item tool - surgically insert a single list item after a specific paragraph
 * @param {number} afterParagraphIndex - 1-based paragraph index to insert after
 * @param {string} text - The text content (without numbering)
 * @param {number} indentLevel - Relative indent: 0=same, 1=deeper, -1=shallower
 */
async function executeInsertListItem(afterParagraphIndex, text, indentLevel = 0) {
  console.log(`[executeInsertListItem] Insert after P${afterParagraphIndex}: "${text.substring(0, 50)}..." (indent: ${indentLevel})`);

  try {
    await Word.run(async (context) => {
      const redlineEnabled = loadRedlineSetting();
      const trackingState = await setChangeTrackingForAi(context, redlineEnabled, "executeInsertListItem");
      try {

        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items/text");
        await context.sync();

        const paraIdx = afterParagraphIndex - 1; // Convert to 0-based
        if (paraIdx < 0 || paraIdx >= paragraphs.items.length) {
          throw new Error(`Paragraph index ${afterParagraphIndex} out of range (1-${paragraphs.items.length})`);
        }

        const adjacentPara = paragraphs.items[paraIdx];

        // Read the adjacent paragraph's OOXML to get its numId and ilvl
        const adjacentOoxml = adjacentPara.getOoxml();
        await context.sync();

        const ooxmlValue = adjacentOoxml.value;
        const numIdMatch = ooxmlValue.match(/<[\w:]*?numId\s+[\w:]*?val="(\d+)"/i);
        const ilvlMatch = ooxmlValue.match(/<[\w:]*?ilvl\s+[\w:]*?val="(\d+)"/i);

        // Debug: Log the numbering definition info if available
        const lvlTextMatch = ooxmlValue.match(/<[\w:]*?lvlText\s+[\w:]*?val="([^"]*)"/i);
        if (lvlTextMatch) {
          console.log(`[executeInsertListItem] Adjacent lvlText format: "${lvlTextMatch[1]}"`);
        }

        // Log a snippet of the OOXML for debugging numbering structure
        const numPrSection = ooxmlValue.match(/<[\w:]*?numPr[\s\S]*?<\/[\w:]*?numPr>/i);
        if (numPrSection) {
          console.log(`[executeInsertListItem] Adjacent numPr: ${numPrSection[0]}`);
        }

        if (!numIdMatch) {
          // Adjacent paragraph is not a list item - just insert plain paragraph
          console.log("[executeInsertListItem] Adjacent paragraph is not a list item, inserting plain paragraph");
          adjacentPara.insertParagraph(text, "After");
          await context.sync();
          return;
        }

        const numId = numIdMatch[1];
        const baseIlvl = ilvlMatch ? parseInt(ilvlMatch[1], 10) : 0;
        const newIlvl = Math.max(0, Math.min(8, baseIlvl + indentLevel)); // Clamp to 0-8

        console.log(`[executeInsertListItem] Adjacent numId=${numId}, ilvl=${baseIlvl}, newIlvl=${newIlvl}`);

        // Extract run properties (rPr) from adjacent paragraph to preserve font styling
        let rPrBlock = '';
        const rPrMatch = ooxmlValue.match(/<[\w:]*?rPr[^>]*>([\s\S]*?)<\/[\w:]*?rPr>/i);
        if (rPrMatch) {
          rPrBlock = rPrMatch[0];
          console.log(`[executeInsertListItem] Extracted rPr from adjacent paragraph`);
        } else {
          const fontMatch = ooxmlValue.match(/<[\w:]*?rFonts[^>]*\/>/i);
          if (fontMatch) {
            rPrBlock = `<w:rPr>${fontMatch[0]}</w:rPr>`;
            console.log(`[executeInsertListItem] Extracted rFonts from adjacent paragraph`);
          }
        }

        // Build OOXML for the new list item
        const escapedText = text
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;');

        const oxmlPara = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
          <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
            <pkg:xmlData>
              <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
              </Relationships>
            </pkg:xmlData>
          </pkg:part>
          <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
            <pkg:xmlData>
              <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:body>
                  <w:p>
                    <w:pPr>
                      <w:pStyle w:val="ListParagraph"/>
                      <w:numPr>
                        <w:ilvl w:val="${newIlvl}"/>
                        <w:numId w:val="${numId}"/>
                      </w:numPr>
                    </w:pPr>
                    <w:r>
                      ${rPrBlock}
                      <w:t xml:space="preserve">${escapedText}</w:t>
                    </w:r>
                  </w:p>
                </w:body>
              </w:document>
            </pkg:xmlData>
          </pkg:part>
        </pkg:package>`;

        // Insert the paragraph with text, then apply list formatting
        const insertedPara = adjacentPara.insertParagraph(text, "After");
        await context.sync();

        // Try to apply the same list formatting using Word's list API
        // The insertedPara should inherit some formatting, but we need to set the list explicitly
        try {
          // Load the inserted paragraph to access its list properties
          insertedPara.load("listItem");
          await context.sync();

          // If it has a listItem, we can adjust its level
          if (insertedPara.listItem && !insertedPara.listItem.isNullObject) {
            // The list item exists - try to adjust level
            console.log(`[executeInsertListItem] Inserted paragraph has listItem, adjusting level to ${newIlvl}`);
            insertedPara.listItem.level = newIlvl;
            await context.sync();
          } else {
            // No listItem - need to add it to a list
            // Use the same numId as adjacent paragraph via OOXML
            console.log(`[executeInsertListItem] No listItem found, applying list via OOXML`);

            const paraRange = insertedPara.getRange("Whole");
            paraRange.insertOoxml(oxmlPara, "Replace");
            await context.sync();
          }
        } catch (listError) {
          console.warn(`[executeInsertListItem] Could not apply list format via API: ${listError.message}`);
          // Fallback: try OOXML replacement
          try {
            const paraRange = insertedPara.getRange("Whole");
            paraRange.insertOoxml(oxmlPara, "Replace");
            await context.sync();
          } catch (oxmlError) {
            console.warn(`[executeInsertListItem] OOXML fallback also failed: ${oxmlError.message}`);
          }
        }

        console.log(`[executeInsertListItem] Successfully inserted list item (numId=${numId}, ilvl=${newIlvl})`);
      } finally {
        await restoreChangeTracking(context, trackingState, "executeInsertListItem");
      }
    });

    return {
      success: true,
      message: `Successfully inserted list item after P${afterParagraphIndex}`
    };
  } catch (error) {
    console.error("[executeInsertListItem] Error:", error);
    return {
      success: false,
      message: `Failed to insert list item: ${error.message}`
    };
  }
}

/**
 * Execute edit_list tool - replaces a range of paragraphs with a proper list
 * Uses the OOXML reconciliation pipeline for list generation and redline fidelity.
 * @param {number} startIndex - 1-based paragraph index of first paragraph
 * @param {number} endIndex - 1-based paragraph index of last paragraph
 * @param {string[]} newItems - Array of new list item texts
 * @param {string} listType - "bullet" or "numbered"
 * @param {string} numberingStyle - For numbered lists: "decimal", "lowerAlpha", "upperAlpha", "lowerRoman", "upperRoman"
 */
async function executeEditList(startIndex, endIndex, newItems, listType, numberingStyle) {
  if (!newItems || newItems.length === 0) {
    return { success: false, message: "No list items provided." };
  }

  console.log(`\n\n========== 📋 EXECUTE_EDIT_LIST CALLED ==========`);
  console.log(`executeEditList: Converting P${startIndex}-P${endIndex} to ${listType} list with ${newItems.length} items`);
  console.log(`[executeEditList] Numbering style: ${numberingStyle}`);
  console.log(`[executeEditList] Raw newItems array:`);
  newItems.forEach((item, idx) => {
    console.log(`  [${idx}]: "${item.substring(0, 60)}${item.length > 60 ? '...' : ''}"`);
  });

  try {
    // Detect and cache document font before entering Word.run.
    await detectDocumentFont();
    let listApplied = false;

    await Word.run(async (context) => {
      const redlineEnabled = loadRedlineSetting();
      const redlineAuthor = loadRedlineAuthor();

      // Disable native tracking while inserting OOXML that already contains
      // explicit w:ins/w:del markup from the reconciliation engine.
      const trackingState = await setChangeTrackingForAi(context, false, "executeEditList");
      try {

        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items/text");
        await context.sync();

        let startIdx = startIndex - 1; // Convert to 0-based
        let endIdx = endIndex - 1;

        // Handle out-of-range paragraph indices gracefully
        // The AI may reference paragraphs that don't exist (e.g., after list expansion)
        const paragraphCount = paragraphs.items.length;

        if (paragraphCount === 0) {
          throw new Error("Document has no paragraphs");
        }

        // If start is beyond document, append at end
        if (startIdx >= paragraphCount) {
          console.log(`Start index ${startIndex} exceeds document (${paragraphCount} paragraphs), treating as append`);
          startIdx = paragraphCount - 1;
          endIdx = paragraphCount - 1;
        }

        // Clamp start to valid range
        if (startIdx < 0) {
          startIdx = 0;
        }

        // Clamp end to valid range
        if (endIdx >= paragraphCount) {
          console.log(`End index ${endIndex} exceeds document (${paragraphCount} paragraphs), clamping to ${paragraphCount}`);
          endIdx = paragraphCount - 1;
        }

        // Ensure start <= end
        if (startIdx > endIdx) {
          startIdx = endIdx;
        }

        console.log(`Adjusted range: P${startIdx + 1} to P${endIdx + 1} (original: ${startIndex} to ${endIndex})`);

        // Get the range covering all paragraphs to replace
        const firstPara = paragraphs.items[startIdx];
        const lastPara = paragraphs.items[endIdx];

        // Get ranges to create a combined range for bulk replacement
        const startRange = firstPara.getRange("Start");
        const endRange = lastPara.getRange("End");
        const fullRange = startRange.expandTo(endRange);
        const rangeOoxmlResult = fullRange.getOoxml();

        await context.sync();

        const normalizedListType = listType === "bullet" ? "bullet" : "numbered";
        const normalizedNumberingStyle = numberingStyle || "decimal";
        console.log(`[executeEditList] Target list format: ${normalizedListType}, numberingStyle: ${normalizedNumberingStyle}`);

        const itemsWithLevels = normalizeListItemsWithLevels(newItems, { indentSpaces: 4 });
        for (const item of itemsWithLevels) {
          if (item.removedMarker) {
            console.log(`[executeEditList] Stripped marker: "${item.removedMarker}" from item`);
          }
          console.log(`[executeEditList] Level: ${item.level}, Text: "${item.text.substring(0, 40)}..."`);
        }

        const existingCount = endIdx - startIdx + 1;
        const newCount = itemsWithLevels.length;

        console.log(`[executeEditList] Replacing ${existingCount} existing paragraphs with ${newCount} list items`);

        const listMarkdown = buildListMarkdown(itemsWithLevels, normalizedListType, normalizedNumberingStyle);
        const originalRangeText = paragraphs.items
          .slice(startIdx, endIdx + 1)
          .map((p) => p.text || "")
          .join("\n");

        console.log(`[executeEditList] Generated list markdown:\n${listMarkdown}`);

        const result = await applyRedlineToOxml(
          rangeOoxmlResult.value,
          originalRangeText,
          listMarkdown,
          {
            author: redlineEnabled ? redlineAuthor : undefined,
            generateRedlines: redlineEnabled
          }
        );

        if (!result?.oxml || !result.hasChanges) {
          console.log("[executeEditList] Reconciliation reported no changes; forcing structural list generation fallback");

          const fallbackPipeline = new ReconciliationPipeline({
            generateRedlines: redlineEnabled,
            author: redlineAuthor
          });

          const fallbackResult = await fallbackPipeline.executeListGeneration(
            listMarkdown,
            null,
            null,
            originalRangeText
          );

          const fallbackOoxml = fallbackResult?.ooxml || fallbackResult?.oxml || "";
          const fallbackIsValid = fallbackResult?.isValid !== false;

          if (!fallbackOoxml || !fallbackIsValid) {
            console.log("[executeEditList] Fallback list generation produced no valid OOXML");
            listApplied = false;
            return;
          }

          const wrappedFallbackOoxml = wrapInDocumentFragment(fallbackOoxml, {
            includeNumbering: true,
            numberingXml: fallbackResult.numberingXml
          });

          fullRange.insertOoxml(wrappedFallbackOoxml, "Replace");
          await context.sync();
          listApplied = true;
          console.log("[executeEditList] Structural fallback replacement succeeded");
          return;
        }

        fullRange.insertOoxml(result.oxml, "Replace");
        await context.sync();
        listApplied = true;
        console.log("[executeEditList] OOXML list reconciliation + replacement succeeded");

      } finally {
        await restoreChangeTracking(context, trackingState, "executeEditList");
      }

      console.log(`\n[executeEditList] ✅ SUCCESSFULLY COMPLETED`);
      console.log(`========== END EXECUTE_EDIT_LIST ==========\n\n`);
      console.log(`Successfully replaced paragraphs with ${listType} list`);
    });

    return {
      success: true,
      message: listApplied
        ? `Successfully created ${listType} list with ${newItems.length} items.`
        : "No list changes were needed."
    };
  } catch (error) {
    console.error("Error in executeEditList:", error);
    return {
      success: false,
      message: `Failed to edit list: ${error.message}`
    };
  }
}

/**
 * Execute convert_headers_to_list tool - converts non-contiguous headers to a numbered list
 * This handles the case where headers like "1. PURPOSE", "2. DEFINITION" have body text between them
 * @param {number[]} paragraphIndices - Array of 1-based paragraph indices of headers to convert
 * @param {string[]} newHeaderTexts - Optional array of new header texts (without numbers)
 * @param {string} numberingFormat - Optional: 'arabic' (default), 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'
 */
/**
 * Execute convert_headers_to_list tool - converts non-contiguous headers to a numbered list
 * This handles the case where headers like "1. PURPOSE", "2. DEFINITION" have body text between them
 * @param {number[]} paragraphIndices - Array of 1-based paragraph indices of headers to convert
 * @param {string[]} newHeaderTexts - Optional array of new header texts (without numbers)
 * @param {string} numberingFormat - Optional: 'arabic' (default), 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'
 */
async function executeConvertHeadersToList(paragraphIndices, newHeaderTexts, numberingFormat) {
  if (!paragraphIndices || paragraphIndices.length === 0) {
    return { success: false, message: "No paragraph indices provided." };
  }

  // Deduplicate paragraph indices to prevent multiple processing
  const distinctIndices = [...new Set(paragraphIndices)];
  if (distinctIndices.length !== paragraphIndices.length) {
    console.log(`Deduplicated indices: ${paragraphIndices.length} -> ${distinctIndices.length}`);
  }

  // Default to arabic if not specified
  const format = numberingFormat || "arabic";
  console.log(`executeConvertHeadersToList: Converting ${distinctIndices.length} headers to ${format} numbered list`);

  try {
    await Word.run(async (context) => {
      const redlineEnabled = loadRedlineSetting();
      const trackingState = await setChangeTrackingForAi(context, redlineEnabled, "executeConvertHeadersToList");
      try {

        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items/text");
        await context.sync();

        // Sort indices to process in order
        const sortedIndices = distinctIndices.sort((a, b) => a - b);

        // Validate all indices
        for (const idx of sortedIndices) {
          const pIdx = idx - 1;
          if (pIdx < 0 || pIdx >= paragraphs.items.length) {
            throw new Error(`Invalid paragraph index: ${idx}`);
          }
        }

        // Get the first header paragraph and start a new list
        const firstIdx = sortedIndices[0] - 1;
        const firstPara = paragraphs.items[firstIdx];
        firstPara.load("text");
        await context.sync();

        // Strip manual numbering from the first header if present
        // Enhanced pattern to catch A., 1.1, I., etc.
        let firstText = firstPara.text || "";
        // Matches: Start of string, optional whitespace, followed by one or more groups of (letters/digits/roman + dot/paren), followed by whitespace
        const numberPattern = /^\s*(?:(?:\d+|[a-zA-Z]+|[ivxlcIVXLC]+)[.)]\s*)+/;
        firstText = firstText.replace(numberPattern, "").trim();

        // Use new text if provided
        if (newHeaderTexts && newHeaderTexts.length > 0) {
          firstText = newHeaderTexts[0];
        }

        // Replace content directly using "Replace" to avoid doubling issues
        // "Replace" overwrites the entire paragraph content cleanly
        firstPara.insertText(firstText, Word.InsertLocation.replace);
        await context.sync();

        // Start a new list on this paragraph
        const list = firstPara.startNewList();
        await context.sync();

        // Load the list to set its numbering format
        list.load("id, levelTypes");
        await context.sync();

        // Map format string to Word.ListNumbering constant
        const numberingMap = {
          "arabic": Word.ListNumbering.arabic,
          "lowerLetter": Word.ListNumbering.lowerLetter,
          "upperLetter": Word.ListNumbering.upperLetter,
          "lowerRoman": Word.ListNumbering.lowerRoman,
          "upperRoman": Word.ListNumbering.upperRoman
        };

        const wordNumbering = numberingMap[format] || Word.ListNumbering.arabic;

        // Set the list to use the specified numbering format
        try {
          list.setLevelNumbering(0, wordNumbering);
          await context.sync();
          console.log(`Set list numbering to ${format}`);
        } catch (numError) {
          console.warn("Could not set level numbering, trying style approach:", numError);
          // Fallback: apply numbered list style
          firstPara.styleBuiltIn = Word.BuiltInStyleName.listNumber;
          await context.sync();
        }

        console.log(`Started new numbered list on paragraph ${sortedIndices[0]}`);

        // OPTIMIZATION: Pre-load text for all remaining headers
        for (let i = 1; i < sortedIndices.length; i++) {
          paragraphs.items[sortedIndices[i] - 1].load("text");
        }
        await context.sync();

        // For remaining headers, attach them to the same list
        for (let i = 1; i < sortedIndices.length; i++) {
          const pIdx = sortedIndices[i] - 1;
          const para = paragraphs.items[pIdx];

          // Strip manual numbering
          let paraText = para.text || "";
          paraText = paraText.replace(numberPattern, "").trim();

          // Use new text if provided (note: using original index mapping could be complex if sorted differently,
          // but assuming 1:1 mapping for sorted newHeaderTexts if they were provided in order of appearance)
          // Ideally newHeaderTexts aligns with the SORTED order if provided by the AI for specific paragraphs.
          // However, usually newHeaderTexts corresponds to input order.
          // For safety, if newHeaderTexts is used, we should map it carefully.
          // IF newHeaderTexts is just a flat list matching the input indices, we might have a mismatch if we sort.
          // But usually this tool is called with indices in document order anyway.

          if (newHeaderTexts && newHeaderTexts.length > i) {
            paraText = newHeaderTexts[i];
          }

          // Replace content directly
          para.insertText(paraText, Word.InsertLocation.replace);
          await context.sync();

          // Attach to the list
          try {
            para.attachToList(list.id, 0); // level 0
            await context.sync();
            console.log(`Attached paragraph ${sortedIndices[i]} to list`);
          } catch (attachError) {
            console.warn(`Could not attach paragraph ${sortedIndices[i]}, using style:`, attachError);
            para.styleBuiltIn = Word.BuiltInStyleName.listNumber;
            await context.sync();
          }
        }

      } finally {
        await restoreChangeTracking(context, trackingState, "executeConvertHeadersToList");
      }

      console.log(`Successfully converted ${distinctIndices.length} headers to numbered list`);
    });

    return {
      success: true,
      message: `Successfully converted ${distinctIndices.length} headers to a numbered list.`
    };
  } catch (error) {
    console.error("Error in executeConvertHeadersToList:", error);
    return {
      success: false,
      message: `Failed to convert headers to list: ${error.message}`
    };
  }
}

/**
 * Execute edit_table tool - performs table operations
 * @param {number} paragraphIndex - 1-based index of any paragraph in the table
 * @param {string} action - "replace_content", "add_row", "delete_row", "update_cell"
 * @param {Array} content - Content for the operation
 * @param {number} targetRow - Target row index (0-based)
 * @param {number} targetColumn - Target column index (0-based)
 */
/**
 * Execute edit_table tool - performs table operations
 * @param {number} paragraphIndex - 1-based index of any paragraph in the table
 * @param {string} action - "replace_content", "add_row", "delete_row", "update_cell"
 * @param {Array} content - Content for the operation
 * @param {number} targetRow - Target row index (0-based)
 * @param {number} targetColumn - Target column index (0-based)
 */
async function executeEditTable(paragraphIndex, action, content, targetRow, targetColumn) {
  try {
    await Word.run(async (context) => {
      const redlineEnabled = loadRedlineSetting();
      const trackingState = await setChangeTrackingForAi(context, redlineEnabled, "executeEditTable");
      let stage = "init";
      try {
        stage = "load_paragraphs";
        const paragraphs = context.document.body.paragraphs;
        // Pre-load text and table relationship
        paragraphs.load("items/text, items/parentTableOrNullObject");
        await context.sync();

        const pIdx = paragraphIndex - 1;
        if (pIdx < 0 || pIdx >= paragraphs.items.length) {
          throw new Error(`Invalid paragraph index: ${paragraphIndex}`);
        }

        const targetPara = paragraphs.items[pIdx];
        if (targetPara.parentTableOrNullObject.isNullObject) {
          throw new Error(`Paragraph ${paragraphIndex} is not inside a table`);
        }

        const table = targetPara.parentTableOrNullObject;
        // Load primary table dimensions first. Rows collection is loaded lazily only when needed.
        stage = "load_table_dimensions";
        table.load("rowCount, columnCount");
        await context.sync();

        const normalizedAction = String(action || "").trim().toLowerCase();

        if (normalizedAction === "replace_content") {
          if (!content || !Array.isArray(content)) {
            throw new Error("replace_content requires a 2D array of content");
          }

          stage = "replace_content_cells";
          let maxRows = Math.min(content.length, table.rowCount || 0);
          if (!(maxRows > 0) && content.length > 0) {
            stage = "load_rows_for_replace_content_fallback";
            table.rows.load("items");
            await context.sync();
            maxRows = Math.min(content.length, table.rows.items.length);
          }
          for (let r = 0; r < maxRows; r++) {
            const rowValues = Array.isArray(content[r]) ? content[r] : [];
            const colsToWrite = rowValues.length;
            for (let c = 0; c < colsToWrite; c++) {
              const value = rowValues[c] == null ? "" : String(rowValues[c]);
              try {
                const cell = table.getCell(r, c);
                // Use replace directly to avoid clear()+start edge cases that can throw ItemNotFound.
                cell.body.insertText(value, Word.InsertLocation.replace);
              } catch (cellError) {
                if (cellError?.code === "ItemNotFound") {
                  console.warn(`[executeEditTable] Skipping missing cell r${r} c${c} during replace_content`);
                  continue;
                }
                throw cellError;
              }
            }
          }
          await context.sync();

        } else if (normalizedAction === "add_row") {
          const rowValues = Array.isArray(content)
            ? (Array.isArray(content[0]) ? content[0] : content)
            : [];
          if (rowValues.length === 0) {
            throw new Error("add_row requires an array of cell values");
          }

          stage = "add_row";
          table.addRows(
            Word.InsertLocation.end,
            1,
            [rowValues.map(value => value == null ? "" : String(value))]
          );
          await context.sync();

        } else if (normalizedAction === "delete_row") {
          stage = "load_rows_for_delete";
          table.rows.load("items");
          await context.sync();

          if (targetRow === undefined || targetRow < 0 || targetRow >= table.rows.items.length) {
            throw new Error(`Invalid or missing row index: ${targetRow}`);
          }

          stage = "delete_row";
          table.rows.items[targetRow].delete();
          await context.sync();

        } else if (normalizedAction === "update_cell") {
          if (targetRow === undefined || targetColumn === undefined || targetRow === null || targetColumn === null) {
            throw new Error("update_cell requires targetRow and targetColumn");
          }
          const parsedTargetRow = Number.parseInt(String(targetRow), 10);
          const parsedTargetColumn = Number.parseInt(String(targetColumn), 10);
          if (!Number.isInteger(parsedTargetRow) || parsedTargetRow < 0) {
            throw new Error(`Invalid row index: ${targetRow}`);
          }
          if (!Number.isInteger(parsedTargetColumn) || parsedTargetColumn < 0) {
            throw new Error(`Invalid column index: ${targetColumn}`);
          }

          let effectiveRowCount = Number.isInteger(table.rowCount) ? table.rowCount : 0;
          let effectiveColumnCount = Number.isInteger(table.columnCount) ? table.columnCount : 0;

          if (!(effectiveRowCount > 0) || !(effectiveColumnCount > 0)) {
            stage = "load_rows_for_update_cell_fallback";
            table.rows.load("items/cellCount");
            await context.sync();

            if (!(effectiveRowCount > 0)) {
              effectiveRowCount = table.rows.items.length;
            }
            if (!(effectiveColumnCount > 0) && parsedTargetRow < table.rows.items.length) {
              const row = table.rows.items[parsedTargetRow];
              effectiveColumnCount = Number.isInteger(row?.cellCount) ? row.cellCount : 0;
            }
          }

          if (effectiveRowCount > 0 && parsedTargetRow >= effectiveRowCount) {
            throw new Error(`Invalid row index: ${parsedTargetRow}`);
          }
          if (effectiveColumnCount > 0 && parsedTargetColumn >= effectiveColumnCount) {
            throw new Error(`Invalid column index: ${parsedTargetColumn}`);
          }

          const cellValue = Array.isArray(content)
            ? (Array.isArray(content[0]) ? content[0][0] : content[0])
            : content;
          const normalizedValue = cellValue == null ? "" : String(cellValue);
          const markdownPreview = preprocessMarkdown(normalizedValue);
          const hasMarkdownFormatting = Array.isArray(markdownPreview?.formatHints) && markdownPreview.formatHints.length > 0;

          stage = "update_cell";
          try {
            const cell = table.getCell(parsedTargetRow, parsedTargetColumn);
            if (hasMarkdownFormatting) {
              stage = "update_cell_shared_markdown_prepare";
              const cellParagraphs = cell.body.paragraphs;
              cellParagraphs.load("items/text");
              await context.sync();

              const currentCellText = cellParagraphs.items
                .map(p => (p.text == null ? "" : String(p.text)))
                .join("\n")
                .trim();

              stage = "update_cell_shared_markdown_apply";
              const applied = await applySharedOperationToWordScope({
                context,
                scope: cell.body.getRange(),
                operation: {
                  type: "redline",
                  targetRef: "P1",
                  target: currentCellText || (markdownPreview.cleanText || normalizedValue),
                  modified: normalizedValue
                },
                author: loadRedlineAuthor(),
                generateRedlines: redlineEnabled,
                logPrefix: "EditTable/Shared"
              });

              if (!applied) {
                throw new Error("Shared markdown cell update produced no changes");
              }
            } else {
              // Use replace directly to avoid clear()+start edge cases.
              cell.body.insertText(normalizedValue, Word.InsertLocation.replace);
            }
          } catch (cellError) {
            if (cellError?.code === "ItemNotFound") {
              throw new Error(`Target cell not found at row ${parsedTargetRow}, column ${parsedTargetColumn}`);
            }
            throw cellError;
          }
          await context.sync();

        } else {
          throw new Error(`Unknown table action: ${action}`);
        }
      } catch (innerError) {
        const codeSuffix = innerError?.code ? ` (${innerError.code})` : "";
        throw new Error(`[executeEditTable/${stage}] ${innerError?.message || innerError}${codeSuffix}`);
      } finally {
        await restoreChangeTracking(context, trackingState, "executeEditTable");
      }
    });

    return {
      success: true,
      message: `Successfully performed table operation: ${action}`
    };
  } catch (error) {
    console.error("Error in executeEditTable:", error);
    return {
      success: false,
      message: `Failed to edit table: ${error.message}`
    };
  }
}

/**
 * Execute edit_section tool - edits a legal contract section
 * @param {number} sectionHeaderIndex - 1-based index of the section header paragraph
 * @param {string} newHeaderText - Optional new text for the header (preserves numbering)
 * @param {string[]} newBodyParagraphs - Optional new body paragraphs
 * @param {boolean} preserveSubsections - Whether to preserve subsections
 */
/**
 * Execute edit_section tool - edits a legal contract section
 * @param {number} sectionHeaderIndex - 1-based index of the section header paragraph
 * @param {string} newHeaderText - Optional new text for the header (preserves numbering)
 * @param {string[]} newBodyParagraphs - Optional new body paragraphs
 * @param {boolean} preserveSubsections - Whether to preserve subsections
 */
async function executeEditSection(sectionHeaderIndex, newHeaderText, newBodyParagraphs, preserveSubsections) {
  try {
    let editCount = 0;

    await Word.run(async (context) => {
      const redlineEnabled = loadRedlineSetting();
      const trackingState = await setChangeTrackingForAi(context, redlineEnabled, "executeEditSection");
      try {

        const paragraphs = context.document.body.paragraphs;
        // OPTIMIZATION: Path-load nested properties to save round-trips
        paragraphs.load("items/text, items/listItemOrNullObject/level");
        await context.sync();

        const headerIdx = sectionHeaderIndex - 1;
        if (headerIdx < 0 || headerIdx >= paragraphs.items.length) {
          throw new Error(`Invalid section header index: ${sectionHeaderIndex}`);
        }

        // Section structure pre-loaded via path syntax above
        const headerPara = paragraphs.items[headerIdx];

        // Check that header is a list item (section header)
        if (headerPara.listItemOrNullObject.isNullObject) {
          throw new Error(`Paragraph ${sectionHeaderIndex} is not a section header (not a list item)`);
        }

        const headerLevel = headerPara.listItemOrNullObject.level || 0;

        // Find the end of this section (next list item at same or higher level)
        let sectionEndIdx = paragraphs.items.length - 1;
        for (let i = headerIdx + 1; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];
          if (!para.listItemOrNullObject.isNullObject) {
            const level = para.listItemOrNullObject.level || 0;
            if (level <= headerLevel) {
              // Found next section at same or higher level
              sectionEndIdx = i - 1;
              break;
            } else if (preserveSubsections) {
              // Found a subsection - stop here if preserving
              sectionEndIdx = i - 1;
              break;
            }
          }
        }

        // Update header text if provided
        if (newHeaderText !== undefined && newHeaderText !== null) {
          // Extract the list number/letter prefix from current text - use robust regex
          const currentText = headerPara.text || "";
          const numberMatch = currentText.match(/^(\d+\.?\s*|\([a-z]\)\s*|[a-z]\.?\s*|[ivxlcdm]+\.?\s*)/i);

          if (numberMatch) {
            // Preserve the numbering prefix
            headerPara.insertText(numberMatch[1] + newHeaderText, Word.InsertLocation.replace);
          } else {
            headerPara.insertText(newHeaderText, Word.InsertLocation.replace);
          }
          editCount++;
        }

        // Replace body paragraphs if provided
        if (newBodyParagraphs && newBodyParagraphs.length > 0) {
          // Delete existing body paragraphs (from end to start)
          for (let i = sectionEndIdx; i > headerIdx; i--) {
            paragraphs.items[i].delete();
          }
          await context.sync();

          // Insert new body paragraphs after header
          let insertAfter = headerPara;
          for (const bodyText of newBodyParagraphs) {
            const newPara = insertAfter.insertParagraph(bodyText, Word.InsertLocation.after);
            insertAfter = newPara;
            editCount++;
          }
        }

        await context.sync();
      } finally {
        await restoreChangeTracking(context, trackingState, "executeEditSection");
      }
    });

    if (editCount === 0) {
      return {
        success: true,
        message: "No changes were specified for the section."
      };
    }

    return {
      success: true,
      message: `Successfully edited section at P${sectionHeaderIndex} (${editCount} changes).`
    };
  } catch (error) {
    console.error("Error in executeEditSection:", error);
    return {
      success: false,
      message: `Failed to edit section: ${error.message}`
    };
  }
}

export {
  initAgenticTools,
  executeRedline,
  executeComment,
  executeHighlight,
  executeNavigate,
  executeResearch,
  executeInsertListItem,
  executeEditList,
  executeConvertHeadersToList,
  executeEditTable,
  executeEditSection
};
