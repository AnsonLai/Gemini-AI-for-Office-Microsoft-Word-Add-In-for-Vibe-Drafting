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
} from '../docx-redline-js-integration/index.js';
import {
  detectDocumentFont
} from '../utils/markdown-utils.js';
import {
  calculateTextStatistics,
  stripFormatting
} from '../utils/text-statistics.js';
import {
  resolveInsertListItemLevel
} from './list-level-utils.js';
import {
  parseAnchoredParagraphs,
  verifyAnchor,
  sanitizeChangeSet,
  formatRejections,
  repairTruncatedJsonArray
} from './change-validation.js';
import {
  getModelProfile
} from '../config/model-profiles.js';
import {
  buildRedlineDiffPrompt,
  buildCorrectiveRetryPrompt,
  REDLINE_DIFF_SCHEMA
} from './redline-prompt.js';

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

export function detectRequestedContentKind(instruction) {
  const normalizedInstruction = String(instruction || '').toLowerCase();
  if (!/\btables?\b/.test(normalizedInstruction)) {
    return null;
  }

  if (/\b(delete|remove|drop)\b.{0,40}\btables?\b|\btables?\b.{0,40}\b(delete|remove|drop)\b|\bwithout (?:a )?table\b|\bno table\b/.test(normalizedInstruction)) {
    return null;
  }

  if (
    /\b(?:turn|convert|make|create|insert|add|format|reformat|restructure|organize|put|place|transform|change)\b[\s\S]{0,120}\b(?:into|as|to|in)\s+(?:a\s+|an\s+)?table\b/.test(normalizedInstruction)
    || /\b(?:into|as)\s+(?:a\s+|an\s+)?table\b/.test(normalizedInstruction)
    || /\btable\b.{0,80}\b(?:to save space|side[- ]by[- ]side|two[- ]column|columns?|rows?)\b/.test(normalizedInstruction)
  ) {
    return 'table';
  }

  return null;
}

async function applyRedlineChangeSet(aiChanges, instruction = '', paragraphTexts = null) {
  const redlineEnabled = loadRedlineSetting();
  const redlineAuthor = loadRedlineAuthor();
  let changesApplied = 0;
  const requestedContentKind = detectRequestedContentKind(instruction);

  // WP1: Verify content anchors before applying. Reject changes whose anchorText
  // does not match the targeted paragraph; auto-correct unambiguous off-by-one
  // indexes (and shift endParagraphIndex by the same offset for replace_range).
  let changesToApply = aiChanges;
  const rejectedChanges = [];
  if (Array.isArray(paragraphTexts) && paragraphTexts.length > 0 && Array.isArray(aiChanges)) {
    changesToApply = [];
    for (const original of aiChanges) {
      const verdict = verifyAnchor(original, paragraphTexts);
      if (!verdict.ok) {
        rejectedChanges.push({
          paragraphIndex: original ? original.paragraphIndex : undefined,
          operation: original ? original.operation : undefined,
          anchorText: original ? original.anchorText : undefined,
          reason: verdict.reason,
          actualTextSnippet: verdict.actualTextSnippet
        });
        continue;
      }

      let change = original;
      if (typeof verdict.correctedIndex === 'number') {
        const offset = verdict.correctedIndex - original.paragraphIndex;
        change = { ...original, paragraphIndex: verdict.correctedIndex };
        if (typeof change.endParagraphIndex === 'number') {
          change.endParagraphIndex = change.endParagraphIndex + offset;
        }
        console.log(`Redline anchor: corrected P${original.paragraphIndex} -> P${verdict.correctedIndex}`);
      }
      changesToApply.push(change);
    }
  }

  if (Array.isArray(changesToApply) && changesToApply.length === 0) {
    return { changesApplied: 0, redlineEnabled, rejectedChanges, engineSkipped: [] };
  }

  let engineSkipped = [];
  await Word.run(async (context) => {
    const trackingState = await setChangeTrackingForAi(context, redlineEnabled, "executeRedline");
    try {
      context.document.load("changeTrackingMode");
      await context.sync();
      const baseTrackingMode = context.document.changeTrackingMode;
      const result = await applyRedlineChangesToWordContext(context, changesToApply, {
        author: redlineAuthor,
        generateRedlines: redlineEnabled,
        disableNativeTracking: redlineEnabled,
        baseTrackingMode,
        logPrefix: "Redline/Shared",
        requestedContentKind
      });
      changesApplied = result.changesApplied;
      engineSkipped = Array.isArray(result.skipped) ? result.skipped : [];
    } finally {
      await restoreChangeTracking(context, trackingState, "executeRedline");
    }
  });

  return {
    changesApplied,
    redlineEnabled,
    rejectedChanges,
    engineSkipped
  };
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

    // 1. Build the prompt for the diff generator (shared with the eval harness)
    const fullPrompt = buildRedlineDiffPrompt(instruction, fullDocumentText);
    const paragraphTexts = parseAnchoredParagraphs(fullDocumentText);

    // Up to 2 attempts: if the first change set is rejected in full (nothing
    // applied, so the document is untouched), re-ask the diff generator ONCE
    // with its own invalid output + the rejection reasons. This fixes malformed
    // change sets (e.g. a thinking model leaving "content" empty and leaking
    // notes into unused fields) without bouncing back to the chat model, which
    // would burn a loop-guard strike on an uninformed retry.
    const MAX_DIFF_ATTEMPTS = 2;
    let lastFailureMessage = null;

    for (let attempt = 1; attempt <= MAX_DIFF_ATTEMPTS; attempt++) {
      const prompt = attempt === 1 || !lastFailureMessage
        ? fullPrompt
        : buildCorrectiveRetryPrompt(fullPrompt, lastFailureMessage.rawChanges, lastFailureMessage.rejectionDetail);

      // 2. Call Gemini to get the JSON array of changes
      const aiChanges = await callGeminiForDiffs(prompt);

      console.log(`AI Suggested Changes (raw, attempt ${attempt}):`, aiChanges);

      if (!aiChanges || !Array.isArray(aiChanges)) {
        lastFailureMessage = {
          rawChanges: aiChanges,
          rejectionDetail: 'The response was not a JSON array.',
          message: "TOOL_FAILURE invalid_response: The diff generator did not return a JSON array. Retry with a simpler instruction or use edit_paragraph operations only."
        };
        continue;
      }

      if (aiChanges.length === 0) {
        // On attempt 1 this means "no edit needed" — a legitimate outcome, not a
        // failure. On the corrective attempt it means the model gave up.
        return {
          message: attempt === 1
            ? "AI had no changes to suggest based on the instruction."
            : `TOOL_FAILURE no_changes_applied: the corrected attempt returned no changes. Original problems:\n${lastFailureMessage?.rejectionDetail || 'unknown'}`,
          showToUser: false
        };
      }

      // WP2: mechanically sanitize the change set (enforce prompt rules in code)
      // before WP1 anchor verification + application.
      const { changes: sanitizedChanges, rejected: sanitizeRejected } =
        sanitizeChangeSet(aiChanges, paragraphTexts.length);

      const { changesApplied, redlineEnabled, rejectedChanges: anchorRejected, engineSkipped } =
        await applyRedlineChangeSet(sanitizedChanges, instruction, paragraphTexts);

      // Merge validation rejections (WP1/WP2) with reasons the apply engine
      // reported for changes it could not apply (e.g. empty target paragraph),
      // so the model gets a concrete reason rather than retrying blindly.
      const allRejected = [...sanitizeRejected, ...anchorRejected, ...(engineSkipped || [])];
      const rejectionDetail = formatRejections(allRejected);

      if (changesApplied === 0) {
        lastFailureMessage = {
          rawChanges: aiChanges,
          rejectionDetail: rejectionDetail || 'No changes could be mapped to the document content.',
          message: rejectionDetail
            ? `TOOL_FAILURE no_changes_applied: ${allRejected.length} change(s) could not be applied.\n${rejectionDetail}\nDo NOT retry the same change. Re-read the document content and target different paragraph(s), fixing the specific issues above.`
            : "Applied 0 edits. The AI's suggestions could not be mapped to the document content."
        };
        if (attempt < MAX_DIFF_ATTEMPTS) {
          console.warn(`Redline attempt ${attempt} applied 0 changes; retrying diff generation with corrective feedback.`);
        }
        continue;
      }

      if (rejectionDetail) {
        const proposed = changesApplied + allRejected.length;
        return {
          message: `Applied ${changesApplied} of ${proposed} edits${redlineEnabled ? ' with redlines' : ' without redlines'}. ${allRejected.length} rejected:\n${rejectionDetail}`,
          showToUser: true
        };
      }

      return {
        message: `Successfully applied ${changesApplied} edits${redlineEnabled ? ' with redlines' : ' without redlines'}.`,
        showToUser: true
      };
    }

    // Both attempts failed; report the last (most informed) failure to the model.
    return {
      message: lastFailureMessage.message,
      showToUser: false
    };

  } catch (error) {
    console.error("Error in executeRedline:", error);
    return {
      message: `Error applying redlines: ${error.message}`,
      showToUser: false  // Silent error - let the model handle it
    };
  }
}
// Hard ceiling for a single diff-generation call. Bounds the damage of model
// repetition loops: a degenerate generation stops burning tokens/time quickly
// instead of grinding to the chat-level 48k budget while the UI hangs.
const DIFF_CALL_TIMEOUT_MS = 90000;

// Helper for the Diff generation (specialized prompt)
async function callGeminiForDiffs(prompt) {
  const geminiApiKey = loadApiKey();
  const geminiModel = loadModel();
  const modelProfile = getModelProfile(geminiModel);
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${geminiModel}:generateContent?key=${geminiApiKey}`;

  const jsonSchema = REDLINE_DIFF_SCHEMA;

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
      temperature: modelProfile.temperature,
      maxOutputTokens: modelProfile.diffMaxOutputTokens || modelProfile.maxOutputTokens,
      responseMimeType: "application/json",
      responseSchema: jsonSchema,
    },
  };

  const abortController = new AbortController();
  const timeoutId = setTimeout(() => abortController.abort(), DIFF_CALL_TIMEOUT_MS);

  try {
    const response = await fetch(apiUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      signal: abortController.signal,
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

    const textPart = candidate.content.parts.find(part => typeof part?.text === "string" && part.text.trim().length > 0);
    if (!textPart) {
      console.error("Gemini diff candidate did not include a text JSON part:", candidate);
      throw new Error("Gemini diff response did not include JSON text.");
    }

    const jsonText = textPart.text;
    console.log("Gemini diff JSON text:", jsonText);

    // Salvage the complete leading objects when the JSON was truncated (e.g. a
    // repetition loop hit maxOutputTokens mid-string). The sanitizer's dedupe
    // then collapses repeated objects, so a degenerate response still yields an
    // actionable change set / rejection reasons instead of a hard parse failure.
    const { changes, repaired } = repairTruncatedJsonArray(jsonText);
    if (repaired) {
      console.warn(`Gemini diff JSON was truncated; salvaged ${changes.length} complete change object(s).`);
    }
    return changes;
  } catch (error) {
    if (error?.name === 'AbortError') {
      console.error(`Diff generation timed out after ${DIFF_CALL_TIMEOUT_MS / 1000}s (likely a model repetition loop).`);
    } else {
      console.error("Error getting diffs:", error);
    }
    return null;
  } finally {
    clearTimeout(timeoutId);
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

    const textPart = candidate.content.parts.find(part => typeof part?.text === "string" && part.text.trim().length > 0);
    if (!textPart) {
      console.error("Gemini JSON candidate did not include a text JSON part:", candidate);
      throw new Error("Gemini JSON response did not include JSON text.");
    }

    const jsonText = textPart.text;
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
        const numPrSection = ooxmlValue.match(/<[\w:]*?numPr[\s\S]*?<\/[\w:]*?numPr>/i);
        const numPrSource = numPrSection ? numPrSection[0] : ooxmlValue;
        const numIdMatch = numPrSource.match(/<[\w:]*?numId\s+[\w:]*?val="(\d+)"/i);
        const ilvlMatch = numPrSource.match(/<[\w:]*?ilvl\s+[\w:]*?val="(\d+)"/i);

        // Debug: Log the numbering definition info if available
        const lvlTextMatch = ooxmlValue.match(/<[\w:]*?lvlText\s+[\w:]*?val="([^"]*)"/i);
        if (lvlTextMatch) {
          console.log(`[executeInsertListItem] Adjacent lvlText format: "${lvlTextMatch[1]}"`);
        }

        // Log a snippet of the OOXML for debugging numbering structure
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
        const levelResolution = resolveInsertListItemLevel(baseIlvl, indentLevel);
        const newIlvl = levelResolution.newIlvl;

        if (levelResolution.appliedIndent !== levelResolution.normalizedIndent) {
          console.warn(
            `[executeInsertListItem] indentLevel=${levelResolution.normalizedIndent} is out of range; clamped to ${levelResolution.appliedIndent}`
          );
        }

        console.log(
          `[executeInsertListItem] Adjacent numId=${numId}, ilvl=${levelResolution.baseIlvl}, indent=${levelResolution.appliedIndent}, newIlvl=${newIlvl}`
        );

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
    return {
      success: false,
      message: `TOOL_FAILURE edit_list: No list items were provided for P${startIndex}-P${endIndex}. Supply the "newItems" array with the list content.`
    };
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
      message: `TOOL_FAILURE edit_list at P${startIndex}-P${endIndex}: ${error.message}`
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
      message: `TOOL_FAILURE edit_table at P${paragraphIndex} (${action}): ${error.message}`
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

/**
 * Host-Agnostic Agentic Tool: Computes word count, character count (with & without spaces),
 * paragraphs, sentences, and reading time for a text selection, text snippet, or paragraph range.
 * Completely free of Office.js / Word / DOM dependencies.
 *
 * @param {object} params - { text, startParagraphIndex, endParagraphIndex }
 * @param {string} fullDocumentText - The full document text with anchors from the chat turn
 * @returns {Promise<{ success: boolean, message: string, stats?: object, scope?: string }>}
 */
async function executeGetSelectionStats({ text, startParagraphIndex, endParagraphIndex } = {}, fullDocumentText = "") {
  try {
    let targetText = "";
    let analyzedScope = "selection";

    // 1. Direct text passed in arguments (highest priority, fully generic)
    if (typeof text === "string" && text.trim().length > 0) {
      targetText = text;
      analyzedScope = "provided text";
    }
    // 2. Paragraph range requested by index
    else if (typeof startParagraphIndex === "number") {
      const startIdx = startParagraphIndex;
      const endIdx = typeof endParagraphIndex === "number" ? endParagraphIndex : startIdx;
      const parsed = parseAnchoredParagraphs(fullDocumentText);

      if (parsed.length > 0) {
        const clampedStart = Math.max(1, Math.min(startIdx, parsed.length));
        const clampedEnd = Math.max(clampedStart, Math.min(endIdx, parsed.length));
        const selectedParagraphs = [];
        for (let i = clampedStart; i <= clampedEnd; i++) {
          const pText = parsed[i - 1];
          if (pText != null) selectedParagraphs.push(pText);
        }
        targetText = selectedParagraphs.join("\n\n");
        analyzedScope = clampedStart === clampedEnd ? `paragraph ${clampedStart}` : `paragraphs ${clampedStart}-${clampedEnd}`;
      }
    }
    // 3. Extract highlighted text from document context string if available
    else if (typeof fullDocumentText === "string" && fullDocumentText.includes("User Highlighted Text:")) {
      const match = fullDocumentText.match(/User Highlighted Text:\s*"""([\s\S]*?)"""/);
      if (match && match[1] && match[1].trim()) {
        targetText = match[1];
        analyzedScope = "selected text";
      }
    }
    // 4. Fallback to full document text if provided
    else if (typeof fullDocumentText === "string" && fullDocumentText.trim()) {
      targetText = fullDocumentText;
      analyzedScope = "document";
    }

    // Check if target text is empty
    if (!targetText || !targetText.trim()) {
      return {
        success: false,
        message: "No text was provided or selected to analyze. Please highlight text or specify a paragraph range.",
        stats: {
          wordCount: 0,
          characterCountWithSpaces: 0,
          characterCountWithoutSpaces: 0,
          paragraphCount: 0,
          sentenceCount: 0,
          estimatedReadingTime: "0 sec"
        }
      };
    }

    const stats = calculateTextStatistics(targetText);

    const message = `Text Statistics (${analyzedScope}):\n`
      + `- Word Count: ${stats.wordCount}\n`
      + `- Character Count (with spaces): ${stats.characterCountWithSpaces}\n`
      + `- Character Count (no spaces): ${stats.characterCountWithoutSpaces}\n`
      + `- Paragraphs: ${stats.paragraphCount}\n`
      + `- Sentences: ${stats.sentenceCount}\n`
      + `- Estimated Reading Time: ${stats.estimatedReadingTime}\n`
      + (stats.preview ? `- Clean Text Preview: "${stats.preview}"` : "");

    return {
      success: true,
      message,
      stats,
      scope: analyzedScope
    };
  } catch (error) {
    console.error("Error in executeGetSelectionStats:", error);
    return {
      success: false,
      message: `Failed to calculate statistics: ${error.message}`
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
  executeEditSection,
  executeGetSelectionStats
};
