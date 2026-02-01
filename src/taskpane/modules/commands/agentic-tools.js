/* global Word */

import { applyRedlineToOxml, ReconciliationPipeline, wrapInDocumentFragment, parseTable, getAuthorForTracking } from '../reconciliation/index.js';
import { applyHighlightToOoxml } from '../../ooxml-formatting-removal.js';
import {
  detectDocumentFont,
  markdownToWordHtml,
  markdownToWordHtmlInline,
  hasBlockElements,
  hasInlineMarkdownFormatting,
  preprocessMarkdownForParagraph,
  applyFormatHintsToRanges,
  applyFormatRemovalToRanges,
  parseMarkdownList,
  normalizeContentEscapes
} from '../utils/markdown-utils.js';

let loadApiKey;
let loadModel;
let loadSystemMessage;
let loadRedlineSetting;
let loadRedlineAuthor;
let setChangeTrackingForAi;
let restoreChangeTracking;
let SEARCH_LIMITS;
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
    SEARCH_LIMITS,
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
- Use "modify_text" ONLY as a fallback for very specific surgical edits where you need to target exact substrings.
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

    let changesApplied = 0;
    const redlineEnabled = loadRedlineSetting();

    // 3. Apply changes in Word
    await Word.run(async (context) => {
      const trackingState = await setChangeTrackingForAi(context, redlineEnabled, "executeRedline");
      try {

        // Load paragraphs with all properties needed by routeChangeOperation to avoid syncs in the loop
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items/text, items/style, items/parentTableCellOrNullObject, items/parentTableOrNullObject");
        await context.sync();

        // Track the current paragraph count (may change as we add/remove paragraphs)
        let currentParagraphCount = paragraphs.items.length;

        for (const change of aiChanges) {
          try {
            console.log("Processing change:", JSON.stringify(change));

            const pIndex = change.paragraphIndex - 1; // 0-based index

            // Check if this is an insertion at the end (index equals or exceeds paragraph count)
            // We're lenient here - any index beyond current count is treated as an append
            const isInsertAtEnd = pIndex >= currentParagraphCount;

            // Only reject negative indices - positive ones that exceed count are handled as appends
            if (pIndex < 0) {
              console.warn(`Invalid paragraph index (negative): ${change.paragraphIndex}`);
              continue;
            }

            // For out-of-bounds indices, reload paragraphs and check again
            if (pIndex >= paragraphs.items.length) {
              // Reload paragraphs collection to get any newly added ones
              // Reload paragraphs collection with all properties to maintain performance
              paragraphs.load("items/text, items/style, items/parentTableCellOrNullObject, items/parentTableOrNullObject");
              await context.sync();
              currentParagraphCount = paragraphs.items.length;

              // If still out of bounds after reload, treat as append to last paragraph
              if (pIndex >= paragraphs.items.length) {
                console.log(`Paragraph index ${change.paragraphIndex} exceeds count (${paragraphs.items.length}), treating as append`);
              }
            }

            // For insertions at the end, use the last paragraph as reference
            const targetParagraph = (pIndex >= paragraphs.items.length)
              ? paragraphs.items[paragraphs.items.length - 1]
              : paragraphs.items[pIndex];

            if (change.operation === "edit_paragraph") {
              console.log(`Editing Paragraph ${change.paragraphIndex} with DMP`);

              if (!change.newContent) {
                console.warn("No newContent provided for edit_paragraph. Skipping.");
                continue;
              }

              try {
                // If inserting at end, insert new paragraph instead of editing
                if (isInsertAtEnd) {
                  console.log(`Inserting new paragraph after paragraph ${paragraphs.items.length}`);
                  targetParagraph.insertParagraph(change.newContent, "After");
                  await context.sync(); // Sync immediately to ensure tracked changes captures the insertion
                  changesApplied++;
                } else {
                  // Route through our smart operation router with preloaded properties
                  await routeChangeOperation(change, targetParagraph, context, true);
                  changesApplied++;
                }
              } catch (error) {
                console.error(`Error editing paragraph ${change.paragraphIndex}:`, error);
                // Fallback to old modify_text approach if DMP fails
              }

            } else if (change.operation === "replace_paragraph") {
              console.log(`Replacing Paragraph ${change.paragraphIndex}`);

              if (change.content === null || change.content === undefined) {
                console.warn("Content is null/undefined for replace_paragraph. Skipping.");
                continue;
              }

              // Normalize content: Convert literal escape sequences to actual characters
              // This handles cases where the AI returns "\\n" as a two-character string instead of actual newlines
              let normalizedContent = normalizeContentEscapes(change.content || "");

              // --- NEW: Detect if target paragraph is already a list item ---
              // If so, we need to preserve its numId/ilvl when replacing content
              let targetIsListItem = false;
              let targetListContext = null;

              if (!isInsertAtEnd) {
                try {
                  const targetOoxmlResult = targetParagraph.getOoxml();
                  await context.sync();

                  // Check for w:numPr in the paragraph's OOXML
                  // Check for w:numPr in the paragraph's OOXML - use robust regex for different XML serializations
                  const numPrMatch = targetOoxmlResult.value.match(/<(?:w:)?numPr>[\s\S]*?<(?:w:)?ilvl\s+(?:w:)?val="(\d+)"[\s\S]*?<(?:w:)?numId\s+(?:w:)?val="(\d+)"[\s\S]*?<\/(?:w:)?numPr>/i);
                  if (numPrMatch) {
                    targetIsListItem = true;
                    targetListContext = {
                      ooxml: targetOoxmlResult.value,
                      ilvl: numPrMatch[1],
                      numId: numPrMatch[2]
                    };
                    console.log(`[replace_paragraph] Target P${change.paragraphIndex} is list item: numId=${targetListContext.numId}, ilvl=${targetListContext.ilvl}`);
                  }
                } catch (ooxmlError) {
                  console.warn("[replace_paragraph] Could not check list context:", ooxmlError);
                }
              }

              const hasHeadingMarkdown = /^\s*#{1,9}\s+/m.test(normalizedContent);
              const hasMarkdownTable = /^\s*\|.*\|/m.test(normalizedContent) && normalizedContent.includes('\n');
              const hasMixedTable = hasMarkdownTable && normalizedContent.replace(/^\s*\|.*$/gm, '').trim().length > 0;

              // If target is a list item and content is plain text (no list markers),
              // use OOXML reconciliation to preserve list formatting
              const contentHasListMarkers = /^(\s*)([-*•]|\d+\.|[a-zA-Z]\.|[ivxlcIVXLC]+\.|\d+\.\d+\.?)\s+/m.test(normalizedContent);
              const contentHasStructuralMarkers = contentHasListMarkers || hasHeadingMarkdown || hasMarkdownTable;
              console.log(`[replace_paragraph] contentHasListMarkers: ${contentHasListMarkers}`);

              if (targetIsListItem && !contentHasStructuralMarkers) {
                console.log(`[replace_paragraph] Preserving list context for plain text edit`);

                try {
                  const redlineEnabled = loadRedlineSetting();
                  const redlineAuthor = loadRedlineAuthor();

                  // Get original text for diffing
                  const originalText = targetParagraph.text;
                  await context.sync();

                  // Use OOXML reconciliation to preserve numPr
                  const result = await applyRedlineToOxml(
                    targetListContext.ooxml,
                    originalText,
                    normalizedContent,
                    {
                      author: redlineEnabled ? redlineAuthor : undefined,
                      generateRedlines: redlineEnabled
                    }
                  );

                  if (result.oxml && result.hasChanges) {
                    const doc = context.document;
                    doc.load("changeTrackingMode");
                    await context.sync();

                    const originalMode = doc.changeTrackingMode;
                    if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                      doc.changeTrackingMode = Word.ChangeTrackingMode.off;
                      await context.sync();
                    }

                    try {
                      targetParagraph.insertOoxml(result.oxml, "Replace");
                      await context.sync();
                      console.log("✅ OOXML list-preserving edit successful");
                      changesApplied++;
                    } finally {
                      if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                        doc.changeTrackingMode = originalMode;
                        await context.sync();
                      }
                    }
                    continue; // Skip other handlers
                  }
                } catch (listPreserveError) {
                  console.warn("[replace_paragraph] List preservation failed, falling back:", listPreserveError);
                  // Fall through to standard handlers
                }
              }
              // --- END NEW ---

              // Check if this is a list or block content with headings/tables - use OOXML pipeline for proper redlines
              const listData = parseMarkdownList(normalizedContent);
              console.log(`[replace_paragraph] listData result: type=${listData?.type}, items=${listData?.items?.length}`);
              const shouldUseBlockPipeline = (listData && listData.type !== 'text') || hasMixedTable;
              if (shouldUseBlockPipeline) {
                const listLabel = listData && listData.type !== 'text' ? listData.type : 'block';
                console.log(`Detected ${listLabel} content in replace_paragraph, using OOXML pipeline`);
                try {
                  // Get original paragraph info for proper diff/redlines and font inheritance
                  // Only get original text if we're REPLACING (not appending)
                  let originalTextForDeletion = '';
                  let paragraphFont = null;
                  if (!isInsertAtEnd) {
                    targetParagraph.load("text,font");
                    await context.sync();
                    originalTextForDeletion = targetParagraph.text;

                    // Get font info for inheritance
                    if (targetParagraph.font) {
                      targetParagraph.font.load("name,size");
                      await context.sync();
                      paragraphFont = targetParagraph.font.name;
                      console.log(`[ListGen] Inheriting font from original paragraph: ${paragraphFont} ${targetParagraph.font.size}pt`);
                    }
                  }

                  // Create reconciliation pipeline with redline settings
                  const redlineEnabled = loadRedlineSetting();
                  const redlineAuthor = loadRedlineAuthor();
                  const pipeline = new ReconciliationPipeline({
                    generateRedlines: redlineEnabled,
                    author: redlineAuthor,
                    font: paragraphFont || 'Calibri' // Inherit font from original paragraph
                  });

                  // Execute block generation (list/table/headings) - this creates OOXML with w:ins/w:del track changes
                  const result = await pipeline.executeListGeneration(
                    normalizedContent,
                    null, // numberingContext - let pipeline determine
                    null, // originalRunModel - not available here
                    originalTextForDeletion // Only pass original text if replacing, not appending
                  );

                  console.log(`[ListGen] Generated ${result.ooxml.length} bytes of OOXML, isInsertAtEnd=${isInsertAtEnd}`);

                  if (result.ooxml && result.isValid) {
                    // Wrap in document fragment for insertOoxml
                    const wrappedOoxml = wrapInDocumentFragment(result.ooxml, {
                      includeNumbering: true,
                      numberingXml: result.numberingXml // Crucial for A, B, C styles
                    });

                    // Temporarily disable Word's track changes to avoid double-tracking
                    // Our w:ins/w:del ARE the track changes
                    const doc = context.document;
                    doc.load("changeTrackingMode");
                    await context.sync();

                    const originalMode = doc.changeTrackingMode;
                    if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                      doc.changeTrackingMode = Word.ChangeTrackingMode.off;
                      await context.sync();
                    }

                    try {
                      // Use 'After' if appending at end, 'Replace' if replacing existing paragraph
                      const insertMode = isInsertAtEnd ? 'After' : 'Replace';
                      console.log(`[ListGen] Using insert mode: ${insertMode}`);
                      targetParagraph.insertOoxml(wrappedOoxml, insertMode);
                      await context.sync();
                      console.log(`✅ OOXML list generation successful`);

                      // TEMP: Spacing workaround disabled - causes GeneralException
                      // Will investigate OOXML structure instead
                      /*
                      // WORKAROUND: Insert a dummy spacing paragraph after the list, then remove it
                      // This forces Word to properly re-evaluate and link the list structure
                      try {
                        // Get all paragraphs to find the newly inserted list items
                        const paragraphs = context.document.body.paragraphs;
                        paragraphs.load("items/text");
                        targetParagraph.load("index");
                        await context.sync();
                        
                        // Find the paragraph at the target index (after replacement/insertion)
                        const targetIdx = targetParagraph.index;
                        
                        // Calculate how many list items were inserted
                        const listItemCount = listData.items.length;
                        
                        // Insert dummy paragraph after the last list item
                        if (targetIdx + listItemCount - 1 < paragraphs.items.length) {
                          const lastListItem = paragraphs.items[targetIdx + listItemCount - 1];
                          const dummyPara = lastListItem.insertParagraph("", "After");
                          await context.sync();
                          
                          console.log(`[ListGen] Inserted dummy spacing paragraph after ${listItemCount} list items`);
                          
                          // Force Word to re-evaluate
                          await context.sync();
                          
                          // TEMP: Leave dummy paragraph to test if it fixes formatting
                          // dummyPara.delete();
                          // await context.sync();
                          
                          console.log(`[ListGen] Left dummy spacing paragraph for testing`);
                        }
                      } catch (spacingError) {
                        console.warn(`[ListGen] Spacing workaround failed (non-critical):`, spacingError.message);
                      }
                      */

                      changesApplied++;
                    } finally {
                      // Restore track changes mode
                      if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                        doc.changeTrackingMode = originalMode;
                        await context.sync();
                      }
                    }
                  } else {
                    console.warn('[ListGen] Pipeline returned invalid result, falling back to HTML');
                    const htmlContent = buildListFallbackHtml(normalizedContent, listData);
                    const insertLocation = isInsertAtEnd ? "After" : "Replace";
                    targetParagraph.insertHtml(htmlContent, insertLocation);
                    changesApplied++;
                  }
                } catch (listError) {
                  console.error(`Error in OOXML list generation:`, listError);
                  // Fallback to HTML if OOXML fails
                  const htmlContent = buildListFallbackHtml(normalizedContent, listData);
                  const insertLocation = isInsertAtEnd ? "After" : "Replace";
                  targetParagraph.insertHtml(htmlContent, insertLocation);
                  changesApplied++;
                }
                // Skip the rest of replace_paragraph handling
                continue;
              }

              // Check if this is a table - use OOXML pipeline
              const matchedTable = normalizedContent.includes('|');
              if (matchedTable) {
                const tableData = parseTable(normalizedContent);
                if (tableData.rows.length > 0 || tableData.headers.length > 0) {
                  console.log(`Detected table in replace_paragraph, using OOXML pipeline`);
                  try {
                    // Create reconciliation pipeline with redline settings
                    const redlineEnabled = loadRedlineSetting();
                    const redlineAuthor = loadRedlineAuthor();
                    const pipeline = new ReconciliationPipeline({
                      generateRedlines: redlineEnabled,
                      author: redlineAuthor
                    });

                    // Execute table generation - this creates OOXML with w:tbl and optional w:ins
                    const result = pipeline.executeTableGeneration(normalizedContent);

                    if (result.ooxml && result.isValid) {
                      // Wrap in document fragment
                      const wrappedOoxml = wrapInDocumentFragment(result.ooxml, {
                        includeNumbering: false
                      });

                      // Disable track changes temporarily
                      const doc = context.document;
                      doc.load("changeTrackingMode");
                      await context.sync();

                      const originalMode = doc.changeTrackingMode;
                      if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                        doc.changeTrackingMode = Word.ChangeTrackingMode.off;
                        await context.sync();
                      }

                      try {
                        const insertMode = isInsertAtEnd ? 'After' : 'Replace';
                        console.log(`[TableGen] Using insert mode: ${insertMode}`);
                        targetParagraph.insertOoxml(wrappedOoxml, insertMode);
                        await context.sync();
                        console.log(`✅ OOXML table generation successful`);
                        changesApplied++;
                      } finally {
                        if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                          doc.changeTrackingMode = originalMode;
                          await context.sync();
                        }
                      }
                    } else {
                      console.warn('[TableGen] Pipeline failed, falling back to HTML');
                      const htmlContent = markdownToWordHtml(normalizedContent);
                      targetParagraph.insertHtml(htmlContent, isInsertAtEnd ? "After" : "Replace");
                      changesApplied++;
                    }
                  } catch (tableError) {
                    console.error(`Error in OOXML table generation:`, tableError);
                    const htmlContent = markdownToWordHtml(normalizedContent);
                    targetParagraph.insertHtml(htmlContent, isInsertAtEnd ? "After" : "Replace");
                    changesApplied++;
                  }
                  // Skip the rest of replace_paragraph handling
                  continue;
                }
              }

              // Convert Markdown to Word-compatible HTML for regular content
              let htmlContent = "";
              try {
                htmlContent = markdownToWordHtml(normalizedContent);
              } catch (markedError) {
                console.error("Error parsing markdown:", markedError);
                htmlContent = normalizedContent; // Fallback to raw text
              }

              // Strip wrapping <p> if present to avoid double paragraphs if Word handles it
              // But only if it's a single simple paragraph (no block elements inside)
              const trimmed = htmlContent.trim();
              const hasSingleParagraph = trimmed.startsWith('<p>') && trimmed.endsWith('</p>') &&
                trimmed.indexOf('</p>', 3) === trimmed.length - 4 &&
                !trimmed.includes('<ul>') && !trimmed.includes('<ol>') &&
                !trimmed.includes('<table') && !trimmed.includes('<h');

              if (hasSingleParagraph) {
                htmlContent = trimmed.substring(3, trimmed.length - 4);
              }

              try {
                // If inserting at end, use insertParagraph to add new content after
                if (isInsertAtEnd) {
                  console.log(`Inserting new paragraph after paragraph ${paragraphs.items.length}`);
                  // Use insertParagraph to add new paragraph after the last one
                  const newPara = targetParagraph.insertParagraph(normalizedContent, "After");
                  await context.sync(); // Sync immediately to ensure tracked changes captures the insertion
                  changesApplied++;
                } else {
                  targetParagraph.insertHtml(htmlContent, "Replace");
                  changesApplied++;
                }
              } catch (wordError) {
                console.error(`Error replacing paragraph ${change.paragraphIndex}:`, wordError);
              }

            } else if (change.operation === "replace_range") {
              const endIndex = change.endParagraphIndex - 1;
              if (endIndex < 0 || endIndex >= paragraphs.items.length || endIndex < pIndex) {
                console.warn(`Invalid end paragraph index: ${change.endParagraphIndex}`);
                continue;
              }

              console.log(`Replacing Range from P${change.paragraphIndex} to P${change.endParagraphIndex}`);

              try {
                const startPara = paragraphs.items[pIndex];
                const endPara = paragraphs.items[endIndex];

                // Check if we are inside a table - wrap in try/catch for safety
                let startHasTable = false;
                let endHasTable = false;
                try {
                  startPara.load("parentTable/id");
                  endPara.load("parentTable/id");
                  await context.sync();
                  startHasTable = !startPara.parentTable.isNullObject;
                  endHasTable = !endPara.parentTable.isNullObject;
                } catch (tableCheckError) {
                  console.warn("Could not check for table context:", tableCheckError);
                  // Continue without table detection
                }

                let targetRange = null;
                let isTableReplacement = false;
                let tableToDelete = null;

                // If both start and end are in the same table
                if (startHasTable && endHasTable) {
                  try {
                    const startTable = startPara.parentTable;
                    const endTable = endPara.parentTable;

                    if (startTable.id === endTable.id) {
                      console.log("Detected same table context. Will replace entire table.");
                      // Strategy: Insert AFTER the table, then delete the table.
                      // This avoids GeneralException when replacing complex structures directly.
                      targetRange = startTable.getRange();
                      isTableReplacement = true;
                      tableToDelete = startTable;
                    } else {
                      console.warn("Start and End paragraphs are in DIFFERENT tables. Falling back to standard range expansion.");
                      targetRange = startPara.getRange().expandTo(endPara.getRange());
                    }
                  } catch (tableError) {
                    console.warn("Error handling table replacement, falling back to range:", tableError);
                    targetRange = startPara.getRange().expandTo(endPara.getRange());
                  }
                } else {
                  // Create a range covering both
                  targetRange = startPara.getRange().expandTo(endPara.getRange());
                }

                // Use 'content' field for replace_range (not replacementText)
                const contentToParse = change.content || change.replacementText || "";

                if (!contentToParse || contentToParse.trim().length === 0) {
                  console.warn("Empty content for replace_range. Skipping.");
                  continue;
                }

                // --- NEW: Detect list structures and use OOXML engine for proper numPr ---
                const hasListMarkers = /^(\s*)([-*•]|\d+\.|[a-zA-Z]\.|[ivxlcIVXLC]+\.|\d+\.\d+\.?)\s+/m.test(contentToParse);

                if (hasListMarkers && !isTableReplacement) {
                  console.log("[replace_range] Detected list markers, using OOXML reconciliation");

                  try {
                    // Get the original text from the range for diffing
                    targetRange.load("text");
                    const originalOoxmlResult = startPara.getOoxml(); // Get OOXML from first paragraph
                    await context.sync();

                    const originalText = targetRange.text || "";
                    const redlineEnabled = loadRedlineSetting();
                    const redlineAuthor = loadRedlineAuthor();

                    // Use the OOXML engine for proper list generation
                    const result = await applyRedlineToOxml(
                      originalOoxmlResult.value,
                      originalText,
                      contentToParse,
                      {
                        author: redlineEnabled ? redlineAuthor : undefined,
                        generateRedlines: redlineEnabled
                      }
                    );

                    if (result.oxml && result.hasChanges) {
                      // Temporarily disable track changes to avoid double-tracking
                      const doc = context.document;
                      doc.load("changeTrackingMode");
                      await context.sync();

                      const originalMode = doc.changeTrackingMode;
                      if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                        doc.changeTrackingMode = Word.ChangeTrackingMode.off;
                        await context.sync();
                      }

                      try {
                        targetRange.insertOoxml(result.oxml, "Replace");
                        await context.sync();
                        changesApplied++;
                        console.log("✅ OOXML list reconciliation successful for replace_range");
                      } finally {
                        if (redlineEnabled && originalMode !== Word.ChangeTrackingMode.off) {
                          doc.changeTrackingMode = originalMode;
                          await context.sync();
                        }
                      }
                      continue; // Skip HTML fallback
                    }
                  } catch (ooxmlError) {
                    console.warn("[replace_range] OOXML reconciliation failed, falling back to HTML:", ooxmlError);
                    // Fall through to HTML path
                  }
                }
                // --- END NEW ---

                // Convert Markdown to Word-compatible HTML (fallback for non-list or table content)
                let htmlContent = "";
                try {
                  htmlContent = markdownToWordHtml(contentToParse);
                } catch (markedError) {
                  console.error("Error parsing markdown for range:", markedError);
                  htmlContent = contentToParse;
                }

                if (isTableReplacement && tableToDelete) {
                  // Insert AFTER the table
                  if (htmlContent && htmlContent.trim().length > 0) {
                    targetRange.insertHtml(htmlContent, "After");
                  }
                  // Delete the old table
                  tableToDelete.delete();
                  changesApplied++;
                } else if (targetRange) {
                  // Standard replacement
                  try {
                    targetRange.insertHtml(htmlContent, "Replace");
                    changesApplied++;
                  } catch (replaceError) {
                    console.warn("Standard insertHtml failed. Trying fallback (Clear + InsertStart).", replaceError);
                    // Fallback: Clear and insert at start
                    try {
                      targetRange.clear(); // Clears content but keeps range
                      targetRange.insertHtml(htmlContent, "Start");
                      changesApplied++;
                    } catch (fallbackError) {
                      console.warn("Fallback (Clear+InsertStart) failed. Trying Nuclear Option (InsertText+InsertHtml).", fallbackError);
                      // Fallback 2: Nuke with text first to reset formatting
                      try {
                        // Replace with a placeholder to reset structure
                        const tempRange = targetRange.insertText(" ", "Replace");
                        tempRange.insertHtml(htmlContent, "Replace");
                        changesApplied++;
                      } catch (nuclearError) {
                        console.error("Replacement failed:", nuclearError);
                      }
                    }
                  }
                }
              } catch (rangeError) {
                console.error(`Error replacing range P${change.paragraphIndex}-P${change.endParagraphIndex}:`, rangeError);
              }
            } else if (change.operation === "modify_text") {
              console.log(`Modifying text in Paragraph ${change.paragraphIndex}: "${change.originalText}" -> "${change.replacementText}"`);

              // Safety check for search string length - Word API has strict limits
              const fullOriginalText = change.originalText;
              if (!fullOriginalText || fullOriginalText.length === 0) {
                console.warn(`Empty search text for modify_text in Paragraph ${change.paragraphIndex}. Skipping.`);
                continue;
              }

              // Word's search API has a practical limit of around 80 characters
              const MAX_SEARCH_LENGTH = 80;
              const needsRangeExpansion = fullOriginalText.length > MAX_SEARCH_LENGTH;
              const searchText = needsRangeExpansion
                ? fullOriginalText.substring(0, MAX_SEARCH_LENGTH)
                : fullOriginalText;

              if (needsRangeExpansion) {
                console.warn(`Search text too long (${fullOriginalText.length} chars), using range expansion strategy.`);
              }

              try {
                // Search ONLY within this paragraph
                const searchResults = targetParagraph.search(searchText, { matchCase: true });
                searchResults.load("items/text");
                await context.sync();

                if (searchResults.items.length > 0) {
                  // Apply to first match only when using range expansion (to avoid ambiguity)
                  const matchesToProcess = needsRangeExpansion ? [searchResults.items[0]] : searchResults.items;

                  for (const item of matchesToProcess) {
                    const replacementText = change.replacementText || "";
                    let htmlReplacement = "";
                    try {
                      // Use inline parsing for modify_text to avoid wrapping in <p> tags
                      // unless the content has block elements
                      htmlReplacement = markdownToWordHtmlInline(replacementText);
                    } catch (markedError) {
                      console.error("Error parsing markdown for modify_text:", markedError);
                      htmlReplacement = replacementText;
                    }

                    // Strip wrapping <p> for simple inline content
                    const trimmed = htmlReplacement.trim();
                    const hasSingleParagraph = trimmed.startsWith('<p>') && trimmed.endsWith('</p>') &&
                      trimmed.indexOf('</p>', 3) === trimmed.length - 4 &&
                      !trimmed.includes('<ul>') && !trimmed.includes('<ol>') &&
                      !trimmed.includes('<table') && !trimmed.includes('<h');

                    if (hasSingleParagraph) {
                      htmlReplacement = trimmed.substring(3, trimmed.length - 4);
                    }

                    try {
                      if (needsRangeExpansion) {
                        // Expand the range to cover the full original text length
                        // Strategy: Find a short suffix from the END of the original text,
                        // then expand the range from prefix start to suffix end
                        const foundRange = item.getRange();

                        try {
                          // Take the LAST 60 chars of the original text as our suffix search
                          // This must be short enough for Word's search API
                          const SUFFIX_LENGTH = 60;
                          const suffixStart = Math.max(0, fullOriginalText.length - SUFFIX_LENGTH);
                          const suffixText = fullOriginalText.substring(suffixStart);

                          console.log(`Range expansion: searching for suffix "${suffixText.substring(0, 30)}..." (${suffixText.length} chars)`);

                          if (suffixText.length >= 5 && suffixText.length <= 80) {
                            const suffixResults = targetParagraph.search(suffixText, { matchCase: true });
                            suffixResults.load("items/text");
                            await context.sync();

                            if (suffixResults.items.length > 0) {
                              // Find the suffix match that comes after our prefix match
                              // by expanding from the found prefix to each suffix candidate
                              let expandedSuccessfully = false;

                              for (const suffixMatch of suffixResults.items) {
                                try {
                                  // Expand from found prefix start to suffix end
                                  const expandedRange = foundRange.expandTo(suffixMatch.getRange("End"));
                                  expandedRange.load("text");
                                  await context.sync();

                                  // Verify the expanded range roughly matches the original length
                                  // Allow some tolerance for whitespace differences
                                  const expandedLength = expandedRange.text.length;
                                  const originalLength = fullOriginalText.length;
                                  const tolerance = Math.max(10, originalLength * 0.1);

                                  if (Math.abs(expandedLength - originalLength) <= tolerance) {
                                    console.log(`Expanded range matches: ${expandedLength} chars (original: ${originalLength})`);
                                    // Use insertHtml with "Replace" for atomic replacement (avoids stale range bug)
                                    expandedRange.insertHtml(htmlReplacement || "", "Replace");
                                    changesApplied++;
                                    expandedSuccessfully = true;
                                    break;
                                  } else {
                                    console.log(`Expanded range length mismatch: ${expandedLength} vs ${originalLength}, trying next suffix match`);
                                  }
                                } catch (expandError) {
                                  console.warn("Could not expand to this suffix match:", expandError.message);
                                }
                              }

                              if (!expandedSuccessfully) {
                                // None of the suffix matches worked, fall back to prefix only
                                console.warn("No valid suffix match found, falling back to prefix-only replacement");
                                // Use insertHtml with "Replace" for atomic replacement
                                item.insertHtml(htmlReplacement || "", "Replace");
                                changesApplied++;
                              }
                            } else {
                              // Suffix not found, fall back to just the found range
                              console.warn("Could not find suffix for range expansion, applying to found range only");
                              // Use insertHtml with "Replace" for atomic replacement
                              item.insertHtml(htmlReplacement || "", "Replace");
                              changesApplied++;
                            }
                          } else {
                            // Suffix invalid length, fall back to just the found range
                            console.warn(`Suffix length invalid (${suffixText.length}), applying to found range only`);
                            // Use insertHtml with "Replace" for atomic replacement
                            item.insertHtml(htmlReplacement || "", "Replace");
                            changesApplied++;
                          }
                        } catch (expandError) {
                          console.warn("Range expansion failed, applying to found range only:", expandError.message);
                          // Use insertHtml with "Replace" for atomic replacement
                          item.insertHtml(htmlReplacement || "", "Replace");
                          changesApplied++;
                        }
                      } else {
                        // Standard case: exact match, delete then insert for clean redline
                        // Use insertHtml with "Replace" for atomic replacement
                        item.insertHtml(htmlReplacement || "", "Replace");
                        changesApplied++;
                      }
                    } catch (modifyError) {
                      console.error("Error applying modify_text:", modifyError);
                    }
                  }
                } else {
                  console.warn(`Could not find text "${searchText}" in Paragraph ${change.paragraphIndex}`);
                }
              } catch (searchError) {
                console.warn(`Search failed for modify_text "${searchText}" in Paragraph ${change.paragraphIndex}:`, searchError.message);

                // Fallback: Try with a shorter search string
                if (searchText.length > 30) {
                  const shorterText = searchText.substring(0, 30);
                  console.log(`Retrying modify_text with shorter search: "${shorterText}"`);
                  try {
                    const retryResults = targetParagraph.search(shorterText, { matchCase: true });
                    retryResults.load("items/text");
                    await context.sync();

                    if (retryResults.items.length > 0) {
                      const replacementText = change.replacementText || "";
                      let htmlReplacement = markdownToWordHtmlInline(replacementText);
                      const trimmed = htmlReplacement.trim();
                      const hasSingleParagraph = trimmed.startsWith('<p>') && trimmed.endsWith('</p>') &&
                        trimmed.indexOf('</p>', 3) === trimmed.length - 4 &&
                        !trimmed.includes('<ul>') && !trimmed.includes('<ol>') &&
                        !trimmed.includes('<table') && !trimmed.includes('<h');

                      if (hasSingleParagraph) {
                        htmlReplacement = trimmed.substring(3, trimmed.length - 4);
                      }
                      // Use insertHtml with "Replace" for atomic replacement
                      retryResults.items[0].insertHtml(htmlReplacement || "", "Replace");
                      changesApplied++;
                    }
                  } catch (retryError) {
                    console.warn(`Retry search also failed for modify_text:`, retryError.message);
                  }
                }
              }
            }

            // Ensure any queued operations for this change are executed here,
            // so errors are caught per-change instead of bubbling as one big GeneralException.
            await context.sync();
          } catch (changeError) {
            console.error("Error applying change:", changeError);
          }
        }

        // Final sync (should usually be a no-op now, but kept for safety)
        await context.sync();
      } finally {
        await restoreChangeTracking(context, trackingState, "executeRedline");
      }
    });

    console.log(`Total changes applied: ${changesApplied} `);

    if (changesApplied === 0) {
      return {
        message: "Applied 0 edits. The AI's suggestions could not be mapped to the document content.",
        showToUser: false  // Silent fallback - don't clutter the log
      };
    }

    return {
      message: `Successfully applied ${changesApplied} edits${redlineEnabled ? ' with redlines' : ' without redlines'}.`,
      showToUser: true
    };

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
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text, items/style");
      await context.sync();

      for (const item of aiComments) {
        const pIndex = item.paragraphIndex - 1;
        if (pIndex < 0 || pIndex >= paragraphs.items.length) continue;

        const targetParagraph = paragraphs.items[pIndex];
        const count = await searchWithFallback(targetParagraph, item.textToFind, context, async (match) => {
          match.insertComment(item.commentContent);
        });
        commentsApplied += count;
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

      // We don't need setChangeTrackingForAi for OOXML highlights as we generate w:rPrChange manually,
      // but if we were using Word API we would.

      try {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items/text, items/style");
        await context.sync();

        for (const item of aiHighlights) {
          const pIndex = item.paragraphIndex - 1;
          if (pIndex < 0 || pIndex >= paragraphs.items.length) continue;

          const targetParagraph = paragraphs.items[pIndex];

          try {
            // Get the paragraph's OOXML
            const paragraphOoxml = targetParagraph.getOoxml();
            await context.sync();

            const originalOoxml = paragraphOoxml.value;
            if (!originalOoxml) {
              console.warn(`Could not get OOXML for paragraph ${item.paragraphIndex}`);
              continue;
            }

            // Apply highlight via pure OOXML manipulation with Redline Support
            const modifiedOoxml = applyHighlightToOoxml(originalOoxml, item.textToFind, normalizedColor, {
              generateRedlines: redlineEnabled,
              author: authorName
            });

            // Only insert if something changed
            if (modifiedOoxml && modifiedOoxml !== originalOoxml) {
              targetParagraph.insertOoxml(modifiedOoxml, Word.InsertLocation.replace);
              await context.sync();
              highlightsApplied++;
              console.log(`[OOXML Highlight] Applied ${normalizedColor} highlight to "${item.textToFind}" in P${item.paragraphIndex}`);
            } else {
              console.warn(`[OOXML Highlight] No matching text found for "${item.textToFind}" in P${item.paragraphIndex}`);
            }
          } catch (highlightError) {
            console.warn(`Failed to highlight "${item.textToFind}":`, highlightError.message);
          }
        }
      } finally {
        // await restoreChangeTracking(context, trackingState, "executeHighlight");
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

/**
 * Searches for text within a paragraph with automatic fallback to shorter text on failure.
 * @param {Word.Paragraph} targetParagraph - The paragraph to search within
 * @param {string} searchText - The text to search for
 * @param {Word.RequestContext} context - Word context for sync operations
 * @param {Function} onSuccess - Callback function to execute on each match (receives match object)
 * @returns {Promise<number>} Number of successful operations
 */
async function searchWithFallback(targetParagraph, searchText, context, onSuccess) {
  let operationsCount = 0;

  // Validate and truncate search text
  if (!searchText || searchText.trim().length === 0) {
    return 0;
  }

  if (searchText.length > SEARCH_LIMITS.MAX_LENGTH) {
    searchText = searchText.substring(0, SEARCH_LIMITS.MAX_LENGTH);
  }

  try {
    const searchResults = targetParagraph.search(searchText, { matchCase: false });
    searchResults.load("items/text");
    await context.sync();

    if (searchResults.items.length > 0) {
      for (const match of searchResults.items) {
        await onSuccess(match);
        operationsCount++;
      }
      return operationsCount;
    }
  } catch (searchError) {
    console.warn(`Search failed for "${searchText}":`, searchError.message);

    // Fallback: Try with shorter text
    if (searchText.length > SEARCH_LIMITS.RETRY_LENGTH) {
      const shorterText = searchText.substring(0, SEARCH_LIMITS.RETRY_LENGTH);
      console.log(`Retrying with shorter search: "${shorterText}"`);

      try {
        const retryResults = targetParagraph.search(shorterText, { matchCase: true });
        retryResults.load("items/text");
        await context.sync();

        if (retryResults.items.length > 0) {
          await onSuccess(retryResults.items[0]);  // Only use first match for fallback
          return 1;
        }
      } catch (retryError) {
        console.warn(`Retry search also failed:`, retryError.message);
      }
    }
  }

  return 0;
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
 * Routes a change operation to the appropriate method
 * Uses native Word APIs for lists/tables, DMP for text edits
 */
async function routeChangeOperation(change, targetParagraph, context, propertiesPreloaded = false) {
  // Properties text, style, parentTableCellOrNullObject, parentTableOrNullObject 
  // should ideally be pre-loaded by the caller (e.g. executeRedline) to avoid syncs here.
  if (!propertiesPreloaded) {
    targetParagraph.load("text, style, parentTableCellOrNullObject, parentTableOrNullObject");
    await context.sync();
  }

  const originalText = targetParagraph.text;
  let newContent = change.newContent || change.content || "";

  // Normalize content: Convert literal escape sequences to actual characters
  // This handles cases where the AI returns "\\n" as a two-character string instead of actual newlines
  newContent = normalizeContentEscapes(newContent);

  // 1. Empty original text - try native APIs first
  if (!originalText || originalText.trim().length === 0) {
    console.log("Empty paragraph detected");

    // Lists are now handled by the OOXML pipeline (Stage 4) for portability

    // Try to parse as table - use OOXML Hybrid Mode even for empty paragraphs
    const matchedTable = newContent.includes('|');
    if (matchedTable) {
      const tableData = parseTable(newContent);
      if (tableData.rows.length > 0 || tableData.headers.length > 0) {
        console.log("Detected table in empty paragraph, using OOXML Hybrid Mode");
        // Fall through to OOXML Engine (Stage 4) which handles empty original text correctly
      }
    } else {
      // ...Existing formatting/HTML check...
    }

    // Check if this is simple text with possible formatting
    // For cells inside tables, prefer OOXML path to avoid nested table issues
    const hasFormatting = hasInlineMarkdownFormatting(newContent);
    if (hasFormatting) {
      console.log("Empty paragraph with formatting - using insertText for simplicity");
      // For empty paragraphs with simple formatting, just insert the text directly
      // The formatting will be applied as markdown symbols which is better than nested tables
      // Use insertText with markdown stripped, then apply formatting separately
      const { cleanText, formatHints } = await preprocessMarkdownForParagraph(newContent);
      targetParagraph.insertText(cleanText, "Replace");
      await context.sync();

      // If there are format hints, apply them using Word's font API
      if (formatHints.length > 0) {
        try {
          await applyFormatHintsToRanges(targetParagraph, cleanText, formatHints, context);
        } catch (formatError) {
          console.warn("Could not apply formatting:", formatError);
        }
      }
      return;
    }

    // Fall back to HTML for other content (no formatting, no tables)
    console.log("Using HTML insertion for empty paragraph");
    const htmlContent = markdownToWordHtml(newContent);
    targetParagraph.insertHtml(htmlContent, "Replace");
    return;
  }

  // 2. Check for structured content types

  // Lists are now handled by the OOXML pipeline (Stage 4) for portability

  // NOTE: Table detection removed here. Let applyRedlineToOxml handle tables 
  // via OOXML Hybrid Mode, which handles both existing tables and text-to-table.

  // 3. Check for block elements (headings, mixed content, etc.)
  if (hasBlockElements(newContent)) {
    console.log("Block elements detected, using HTML replacement");
    const htmlContent = markdownToWordHtml(newContent);
    targetParagraph.insertHtml(htmlContent, "Replace");
    return;
  }

  // 4. Use OOXML Engine V5.1 (Hybrid Mode) for proper track changes
  // This modifies the DOM in-place, embedding w:ins/w:del directly in the structure
  console.log("Attempting OOXML Hybrid Mode for text edit");
  const redlineEnabled = loadRedlineSetting();

  // Get original text and paragraph OOXML
  if (!propertiesPreloaded) {
    targetParagraph.load("text");
    await context.sync();
  }

  const paragraphOriginalText = targetParagraph.text;
  let paragraphOoxmlResult = null;
  let paragraphOoxmlValue = null;
  try {
    paragraphOoxmlResult = targetParagraph.getOoxml();
    await context.sync();
    paragraphOoxmlValue = paragraphOoxmlResult.value;
  } catch (ooxmlError) {
    console.warn("[OxmlEngine] Paragraph.getOoxml failed, trying range.getOoxml", ooxmlError);
    try {
      const paragraphRange = targetParagraph.getRange();
      paragraphOoxmlResult = paragraphRange.getOoxml();
      await context.sync();
      paragraphOoxmlValue = paragraphOoxmlResult.value;
    } catch (rangeError) {
      console.warn("[OxmlEngine] Range.getOoxml failed for paragraph", rangeError);
      // Try parent table cell or table as a last resort (pure OOXML path)
      try {
        if (!propertiesPreloaded) {
          targetParagraph.load("parentTableCellOrNullObject, parentTableOrNullObject");
          await context.sync();
        }

        if (targetParagraph.parentTableCellOrNullObject && !targetParagraph.parentTableCellOrNullObject.isNullObject) {
          try {
            console.warn("[OxmlEngine] Trying parent table cell getOoxml");
            paragraphOoxmlResult = targetParagraph.parentTableCellOrNullObject.getOoxml();
            await context.sync();
            paragraphOoxmlValue = paragraphOoxmlResult.value;
          } catch (cellError) {
            console.warn("[OxmlEngine] Parent table cell getOoxml failed, trying cell range", cellError);
            try {
              const cellRange = targetParagraph.parentTableCellOrNullObject.getRange();
              paragraphOoxmlResult = cellRange.getOoxml();
              await context.sync();
              paragraphOoxmlValue = paragraphOoxmlResult.value;
            } catch (cellRangeError) {
              console.warn("[OxmlEngine] Parent table cell range getOoxml failed", cellRangeError);
            }
          }
        }

        if (!paragraphOoxmlValue && targetParagraph.parentTableOrNullObject && !targetParagraph.parentTableOrNullObject.isNullObject) {
          try {
            console.warn("[OxmlEngine] Trying parent table getOoxml");
            paragraphOoxmlResult = targetParagraph.parentTableOrNullObject.getOoxml();
            await context.sync();
            paragraphOoxmlValue = paragraphOoxmlResult.value;
          } catch (tableError) {
            console.warn("[OxmlEngine] Parent table getOoxml failed, trying table range", tableError);
            try {
              const tableRange = targetParagraph.parentTableOrNullObject.getRange();
              paragraphOoxmlResult = tableRange.getOoxml();
              await context.sync();
              paragraphOoxmlValue = paragraphOoxmlResult.value;
            } catch (tableRangeError) {
              console.warn("[OxmlEngine] Parent table range getOoxml failed", tableRangeError);
            }
          }
        }
      } catch (tableError) {
        console.warn("[OxmlEngine] Table OOXML fallback failed", tableError);
      }
    }
  }

  if (!paragraphOoxmlValue) {
    console.warn("[OxmlEngine] Unable to retrieve OOXML for paragraph; skipping OOXML edit");
    return;
  }

  console.log("[OxmlEngine] Original text:", paragraphOriginalText.length > 500 ? paragraphOriginalText.substring(0, 500) + "..." : paragraphOriginalText);
  console.log("[OxmlEngine] Original text length:", paragraphOriginalText.length);

  // Apply redlines using hybrid engine (DOM manipulation approach)
  const redlineAuthor = loadRedlineAuthor();
  const result = await applyRedlineToOxml(
    paragraphOoxmlValue,
    paragraphOriginalText,
    newContent,
    {
      author: redlineEnabled ? redlineAuthor : undefined,
      generateRedlines: redlineEnabled
    }
  );

  if (!result.hasChanges) {
    console.log("[OxmlEngine] No changes detected by engine");
    return;
  }

  // Handle native API formatting for table cells (format addition)
  if (result.useNativeApi && result.formatHints) {
    console.log("[OxmlEngine] Using native Font API for table cell formatting");
    await applyFormatHintsToRanges(targetParagraph, result.originalText, result.formatHints, context);
    console.log("✅ Native API formatting successful");
    return;
  }

  // Handle SURGICAL format changes (pure OOXML at range level, not paragraph level)
  // This searches for specific text and replaces just that range with OOXML
  if (result.isSurgicalFormatChange && result.surgicalChanges) {
    console.log(`[OxmlEngine] Applying ${result.surgicalChanges.length} surgical format changes`);

    const doc = context.document;
    doc.load("changeTrackingMode");
    await context.sync();

    const originalMode = doc.changeTrackingMode;

    // Disable track changes - our OOXML has embedded w:del/w:ins
    if (originalMode !== Word.ChangeTrackingMode.off) {
      console.log("[OxmlEngine] Temporarily disabling track changes for surgical OOXML");
      doc.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
    }

    let successfulSurgicalChanges = 0;

    try {
      const paragraphRange = targetParagraph.getRange();
      paragraphRange.load('text');
      await context.sync();

      for (const change of result.surgicalChanges) {
        try {
          console.log(`[OxmlEngine] Surgical: searching for "${change.searchText}"`);

          // Search for the specific text within the paragraph
          const searchResults = paragraphRange.search(change.searchText, {
            matchCase: true,
            matchWholeWord: false
          });
          searchResults.load('items/text');
          await context.sync();

          if (searchResults.items.length > 0) {
            // Replace at range level - NOT paragraph level
            const targetRange = searchResults.items[0];
            targetRange.insertOoxml(change.replacementOoxml, 'Replace');
            await context.sync();
            console.log(`[OxmlEngine] ✅ Surgical replacement applied for "${change.searchText}"`);
            successfulSurgicalChanges++;
          } else {
            console.warn(`[OxmlEngine] Text not found for surgical replacement: "${change.searchText}"`);
          }
        } catch (changeError) {
          console.warn(`[OxmlEngine] Failed to apply surgical change: ${changeError.message}`);
        }
      }

      if (successfulSurgicalChanges === result.surgicalChanges.length) {
        console.log("✅ Surgical format changes completed");
      } else {
        console.warn(`[OxmlEngine] Surgical replacements applied: ${successfulSurgicalChanges}/${result.surgicalChanges.length}`);
      }
    } finally {
      // Restore track changes mode
      if (originalMode !== Word.ChangeTrackingMode.off) {
        doc.changeTrackingMode = originalMode;
        await context.sync();
      }
    }

    if (successfulSurgicalChanges === result.surgicalChanges.length) {
      return;
    }

    console.warn("[OxmlEngine] Surgical format removal incomplete; no fallback configured");
    return;
  }

  // Handle native API format REMOVAL (e.g., unbold, unitalicize)
  // This uses Word's Font API which properly tracks format changes
  if (result.useNativeApi && result.formatRemovalHints) {
    console.log("[OxmlEngine] Using native Font API for format removal");
    await applyFormatRemovalToRanges(targetParagraph, result.originalText, result.formatRemovalHints, context);
    console.log("✅ Native API format removal successful");
    return;
  }

  console.log("[OxmlEngine] Generated OOXML with track changes, length:", result.oxml.length);

  try {
    // The hybrid engine embeds w:ins/w:del directly in the DOM structure
    // For TEXT changes (w:ins/w:del), we disable Word's track changes to prevent double-tracking
    // For FORMAT-ONLY changes (w:rPrChange), we KEEP track changes ON so Word surfaces our markers
    const doc = context.document;
    doc.load("changeTrackingMode");
    await context.sync();

    const originalMode = doc.changeTrackingMode;
    const shouldDisableTracking = !result.isFormatOnly && originalMode !== Word.ChangeTrackingMode.off;
    console.log(`[OxmlEngine] Current track changes mode: ${originalMode}, redlineEnabled: ${redlineEnabled}, isFormatOnly: ${result.isFormatOnly}, shouldDisableTracking: ${shouldDisableTracking}`);

    // Only disable track changes for TEXT changes (with w:ins/w:del)
    // Keep track changes ON for format-only changes so Word surfaces w:rPrChange
    if (shouldDisableTracking) {
      console.log("[OxmlEngine] Temporarily disabling Word track changes for text-based OOXML insertion");
      doc.changeTrackingMode = Word.ChangeTrackingMode.off;
      await context.sync();
    }

    try {
      // Insert the modified OOXML - since it's a paragraph-level replacement,
      // and our DOM already contains the track change markers embedded in the structure,
      // Word should render them as track changes
      targetParagraph.insertOoxml(result.oxml, 'Replace');
      await context.sync();
      console.log("✅ OOXML Hybrid Mode reconciliation successful");

      // WORKAROUND: If this was a list transformation, insert dummy paragraph to force Word re-evaluation
      // Detect if the result contains list formatting
      if (result.oxml.includes('<w:numPr>') || result.oxml.includes('ListParagraph')) {
        try {
          console.log('[OxmlEngine] Detected list in result, applying spacing workaround');

          // Count how many paragraphs were generated (count <w:p> tags)
          const pCount = (result.oxml.match(/<w:p>/g) || []).length;

          if (pCount > 1) {
            // Reload paragraphs to get the newly inserted ones
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items/text");
            await context.sync();

            // Find the target paragraph index
            const targetIdx = targetParagraph.index || 0;

            // Insert dummy paragraph after the last list item
            if (targetIdx + pCount - 1 < paragraphs.items.length) {
              const lastListItem = paragraphs.items[targetIdx + pCount - 1];
              const dummyPara = lastListItem.insertParagraph("", "After");
              await context.sync();

              console.log(`[OxmlEngine] Inserted dummy spacing paragraph after ${pCount} list items`);

              // Force Word to re-evaluate
              await context.sync();

              // TEMP: Leave dummy paragraph to test if it fixes formatting
              // dummyPara.delete();
              // await context.sync();

              console.log(`[OxmlEngine] Left dummy spacing paragraph for testing`);
            }
          }
        } catch (spacingError) {
          console.warn(`[OxmlEngine] Spacing workaround failed (non-critical):`, spacingError.message);
        }
      }
    } finally {
      // Restore track changes mode (only if we disabled it)
      if (shouldDisableTracking) {
        console.log(`[OxmlEngine] Restoring track changes mode to: ${originalMode}`);
        doc.changeTrackingMode = originalMode;
        await context.sync();
      }
    }
  } catch (insertError) {
    console.error("❌ OOXML insertion failed:", insertError.message);
    // Fallback to simple text replacement
    console.log("Falling back to simple text replacement");
    targetParagraph.insertText(newContent, "Replace");
    await context.sync();
  }
}

function buildListFallbackHtml(normalizedContent, listData) {
  const listHtml = buildHtmlFromListData(listData);
  if (listHtml) {
    return markdownToWordHtml(listHtml);
  }
  return markdownToWordHtml(normalizedContent);
}

function buildHtmlFromListData(listData) {
  if (!listData || !Array.isArray(listData.items) || listData.items.length === 0) {
    return "";
  }

  const root = { type: "root", children: [] };
  const listStack = [];

  const renderInline = (text) => markdownToWordHtmlInline(text || "");

  for (const item of listData.items) {
    if (item.type === "text") {
      listStack.length = 0;
      if (item.text && item.text.trim().length > 0) {
        root.children.push({ type: "p", html: renderInline(item.text) });
      }
      continue;
    }

    const level = Math.max(0, item.level || 0);
    while (listStack.length > level) {
      listStack.pop();
    }

    const parent = level === 0 ? root : (listStack[level - 1]?.lastItem || root);
    if (!parent.children) parent.children = [];

    const tag = item.type === "bullet" ? "ul" : "ol";
    const styleType = item.type === "bullet"
      ? getBulletListStyle(level)
      : getNumberedListStyle(item.marker);

    let listNode = parent.children[parent.children.length - 1];
    if (!listNode || listNode.type !== "list" || listNode.tag !== tag || listNode.styleType !== styleType) {
      listNode = { type: "list", tag, styleType, items: [] };
      parent.children.push(listNode);
    }

    const listItem = { type: "li", html: renderInline(item.text || ""), children: [] };
    listNode.items.push(listItem);

    listStack.length = level;
    listStack[level] = { listNode, lastItem: listItem };
  }

  return renderNodes(root.children);
}

function renderNodes(nodes) {
  if (!nodes || nodes.length === 0) return "";
  return nodes.map(renderNode).join("");
}

function renderNode(node) {
  if (!node) return "";
  if (node.type === "p") {
    return `<p>${node.html}</p>`;
  }
  if (node.type === "list") {
    const style = `style="list-style-type: ${node.styleType}; margin-left: 0; padding-left: 40px; margin-bottom: 10px;"`;
    const items = node.items.map(renderListItem).join("");
    return `<${node.tag} ${style}>${items}</${node.tag}>`;
  }
  return "";
}

function renderListItem(item) {
  const children = renderNodes(item.children);
  return `<li style="margin-bottom: 5px;">${item.html}${children}</li>`;
}

function getBulletListStyle(level) {
  const styles = ["disc", "circle", "square"];
  return styles[level % styles.length];
}

function getNumberedListStyle(marker) {
  const raw = (marker || "").trim();
  if (!raw) return "decimal";

  const cleaned = raw.replace(/^\(|\)$/g, "").replace(/\.$/, "").trim();
  if (/^\d+(\.\d+)*$/.test(cleaned)) return "decimal";
  if (/^[A-Z]$/.test(cleaned)) return "upper-alpha";
  if (/^[a-z]$/.test(cleaned)) return "lower-alpha";
  if (/^[IVXLCDM]+$/.test(cleaned)) return "upper-roman";
  if (/^[ivxlcdm]+$/.test(cleaned)) return "lower-roman";
  return "decimal";
}

/**
 * Fallback function for modify_text operations
 * Used when DMP approach fails
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
 * Uses HTML insertion for reliable list formatting
 * @param {number} startIndex - 1-based paragraph index of first paragraph
 * @param {number} endIndex - 1-based paragraph index of last paragraph
 * @param {string[]} newItems - Array of new list item texts
 * @param {string} listType - "bullet" or "numbered"
 * @param {string} numberingStyle - For numbered lists: "decimal", "lowerAlpha", "upperAlpha", "lowerRoman", "upperRoman"
 */
/**
 * Execute edit_list tool - replaces a range of paragraphs with a proper list
 * Uses HTML insertion for reliable list formatting
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
    await Word.run(async (context) => {
      // Detect document font for consistent HTML insertion
      await detectDocumentFont();

      const redlineEnabled = loadRedlineSetting();
      const trackingState = await setChangeTrackingForAi(context, redlineEnabled, "executeEditList");
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

        // Pre-process: Split paragraphs with soft breaks (\v) if we have more new items than paragraphs
        // This prevents deleting a multi-line paragraph just to re-insert it as list items
        if (newItems.length > (endIdx - startIdx + 1)) {
          let splitsPerformed = 0;

          // Check relevant paragraphs for soft breaks
          const initialEndIdx = endIdx;
          for (let i = startIdx; i <= initialEndIdx; i++) {
            const para = paragraphs.items[i];
            para.load("text");
          }
          await context.sync();

          // We iterate through the originally identified paragraphs
          // Note: references to paragraphs.items[i] remain valid for operation even if others split?
          // Actually, let's grab the object references safely
          const parasToCheck = [];
          for (let i = startIdx; i <= initialEndIdx; i++) {
            parasToCheck.push(paragraphs.items[i]);
          }

          for (const para of parasToCheck) {
            if (para.text && para.text.includes('\u000b')) {
              console.log(`[executeEditList] Found soft breaks in matched paragraph, splitting...`);
              const ranges = para.search("\u000b"); // Search for vertical tab
              ranges.load("items/text");
              await context.sync();

              let paraSplits = 0;
              // Iterate backwards to keep ranges valid during modification
              for (let j = ranges.items.length - 1; j >= 0; j--) {
                ranges.items[j].insertParagraph("", "After");
                ranges.items[j].delete();
                paraSplits++;
              }

              if (paraSplits > 0) {
                await context.sync(); // Commit the splits for this paragraph
                splitsPerformed += paraSplits;
              }
            }
          }

          if (splitsPerformed > 0) {
            console.log(`[executeEditList] Performed ${splitsPerformed} splits. Reloading paragraphs.`);
            endIdx += splitsPerformed; // Extend the range to include new paragraphs

            // Reload paragraphs for the main logic as indices have shifted
            paragraphs.load("items/text");
            await context.sync();
          }
        }

        // Get the range covering all paragraphs to replace
        const firstPara = paragraphs.items[startIdx];
        const lastPara = paragraphs.items[endIdx];

        // Try to read the existing list's numId from the first paragraph's OOXML
        let existingNumId = null;
        let existingBaseIlvl = 0;
        try {
          const firstParaOoxml = firstPara.getOoxml();
          await context.sync();

          const numIdMatch = firstParaOoxml.value.match(/w:numId w:val="(\d+)"/);
          const ilvlMatch = firstParaOoxml.value.match(/w:ilvl w:val="(\d+)"/);

          if (numIdMatch) {
            existingNumId = numIdMatch[1];
            existingBaseIlvl = ilvlMatch ? parseInt(ilvlMatch[1], 10) : 0;
            console.log(`[executeEditList] Found existing numId: ${existingNumId}, base ilvl: ${existingBaseIlvl}`);
          }
        } catch (oxmlError) {
          console.warn(`[executeEditList] Could not read existing OOXML:`, oxmlError.message);
        }

        // Get ranges to create a combined range
        const startRange = firstPara.getRange("Start");
        const endRange = lastPara.getRange("End");
        const fullRange = startRange.expandTo(endRange);

        await context.sync();

        // Build HTML list
        const listTag = listType === "numbered" ? "ol" : "ul";

        // Map numbering style to CSS list-style-type
        let cssListStyleType = "disc"; // default for bullet
        if (listType === "numbered") {
          const styleMap = {
            "decimal": "decimal",
            "lowerAlpha": "lower-alpha",
            "upperAlpha": "upper-alpha",
            "lowerRoman": "lower-roman",
            "upperRoman": "upper-roman"
          };
          cssListStyleType = styleMap[numberingStyle] || "decimal";
        }

        const listStyle = `style="list-style-type: ${cssListStyleType}; margin-left: 0; padding-left: 40px;"`;

        // Map numberingStyle to OOXML numFmt values for direct OOXML insertion
        const numFmtMap = {
          "decimal": "decimal",
          "lowerAlpha": "lowerLetter",
          "upperAlpha": "upperLetter",
          "lowerRoman": "lowerRoman",
          "upperRoman": "upperRoman"
        };
        const numFmt = numFmtMap[numberingStyle] || "decimal";

        // Determine template numId when creating a new list (no existing numId)
        // We'll use numId 100+ for custom styles to avoid conflicts
        let templateNumId = existingNumId;
        if (!templateNumId) {
          if (listType === "bullet") {
            templateNumId = "1"; // Default bullet
          } else {
            // For numbered lists, use numId 2 (default decimal) - the numFmt will override display
            templateNumId = "2";
          }
          console.log(`[executeEditList] No existing numId, using template: ${templateNumId}, numFmt: ${numFmt}`);
        }
        // Detect hierarchy from leading whitespace indentation (4 spaces = 1 level)
        // Also strip any leading list markers from items to avoid doubled numbering
        const markersRegex = /^((?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|[-*•])\s*)/;

        // Analyze items for hierarchy based on leading whitespace
        const itemsWithLevels = newItems.map(item => {
          // Count leading spaces/tabs
          const indentMatch = item.match(/^(\s*)/);
          const indentSize = indentMatch ? indentMatch[1].length : 0;
          const level = Math.floor(indentSize / 4); // 4 spaces per level

          // Strip leading whitespace
          let stripped = item.trim();

          // Also strip any list markers (1., a., -, etc.)
          const markerMatch = stripped.match(markersRegex);
          if (markerMatch) {
            stripped = stripped.replace(markersRegex, '');
            console.log(`[executeEditList] Stripped marker: "${markerMatch[1].trim()}" from item`);
          }

          console.log(`[executeEditList] Level: ${level}, Text: "${stripped.substring(0, 40)}..."`);

          return { text: stripped.trim(), level };
        });

        // SURGICAL APPROACH: Edit existing paragraphs in place
        // This preserves the document's existing formatting better than bulk replacement
        const existingCount = endIdx - startIdx + 1;
        const newCount = itemsWithLevels.length;

        console.log(`[executeEditList] Surgical mode: ${existingCount} existing → ${newCount} new items`);

        // PHASE 1: Edit existing paragraphs with new text (keeping their style)
        const editLimit = Math.min(existingCount, newCount);

        // OPTIMIZATION: Pre-load text and OOXML for all paragraphs to avoid per-paragraph syncs
        const ooxmlResults = [];
        for (let i = 0; i < editLimit; i++) {
          const para = paragraphs.items[startIdx + i];
          para.load("text");
          ooxmlResults.push(para.getOoxml());
        }
        await context.sync();

        for (let i = 0; i < editLimit; i++) {
          const para = paragraphs.items[startIdx + i];
          const item = itemsWithLevels[i];

          const originalText = para.text.trim();
          console.log(`[executeEditList] P${startIdx + i + 1} BEFORE: "${originalText.substring(0, 50)}..."`);
          console.log(`[executeEditList] P${startIdx + i + 1} NEW: "${item.text.substring(0, 50)}..."`);

          const targetOoxmlValue = ooxmlResults[i].value;
          const numIdMatch = targetOoxmlValue.match(/w:numId\s+w:val="(\d+)"/i);
          const ilvlMatch = targetOoxmlValue.match(/w:ilvl\s+w:val="(\d+)"/i);
          const currentNumId = numIdMatch ? numIdMatch[1] : (existingNumId || '1');
          const currentIlvl = ilvlMatch ? ilvlMatch[1] : '0';
          const newIlvl = existingBaseIlvl + item.level;

          console.log(`[executeEditList] P${startIdx + i + 1} numId=${currentNumId}, ilvl: ${currentIlvl} → ${newIlvl}`);

          // Build OOXML that preserves the paragraph's numbering but updates text and ilvl
          const escapedText = item.text
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');

          // Use templateNumId when paragraph doesn't have existing numbering
          const numIdForPara = currentNumId === (existingNumId || '1') && !existingNumId ? templateNumId : currentNumId;

          // Build the numbering.xml part if creating a new list with custom style
          let numberingPart = '';
          if (!existingNumId && listType === "numbered" && numFmt !== 'decimal') {
            // Include a complete numbering definition for custom styles
            // Use a high numId to avoid conflicts (100+)
            const customNumId = '100';
            numberingPart = `
            <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
              <pkg:xmlData>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
              </pkg:xmlData>
            </pkg:part>
            <pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">
              <pkg:xmlData>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="100">
                    <w:multiLevelType w:val="multilevel"/>
                    <w:lvl w:ilvl="0">
                      <w:start w:val="1"/>
                      <w:numFmt w:val="${numFmt}"/>
                      <w:lvlText w:val="%1."/>
                      <w:lvlJc w:val="left"/>
                      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="100">
                    <w:abstractNumId w:val="100"/>
                  </w:num>
                </w:numbering>
              </pkg:xmlData>
            </pkg:part>`;
            // Use this custom numId for the paragraph
          }

          const oxmlPara = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
            <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
              <pkg:xmlData>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
              </pkg:xmlData>
            </pkg:part>${numberingPart}
            <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
              <pkg:xmlData>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val="ListParagraph"/>
                        <w:numPr>
                          <w:ilvl w:val="${newIlvl}"/>
                          <w:numId w:val="${numberingPart ? '100' : numIdForPara}"/>
                        </w:numPr>
                      </w:pPr>
                      <w:r>
                        <w:t xml:space="preserve">${escapedText}</w:t>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
              </pkg:xmlData>
            </pkg:part>
          </pkg:package>`;

          // Use the paragraph's range for replacement (not the paragraph object)
          const paraRange = para.getRange("Whole");
          paraRange.insertOoxml(oxmlPara, "Replace");
          console.log(`[executeEditList] Replaced P${startIdx + i + 1} range with OOXML`);
        }
        await context.sync();

        // PHASE 2: Insert new paragraphs if more items than existing
        if (newCount > existingCount) {
          console.log(`[executeEditList] Phase 2: Inserting ${newCount - existingCount} new paragraphs`);

          // Reload paragraphs after Phase 1 edits
          paragraphs.load("items/text");
          await context.sync();

          // Get the last edited paragraph to insert after
          const lastEditedIdx = startIdx + existingCount - 1;
          const insertAfterPara = paragraphs.items[lastEditedIdx];

          console.log(`[executeEditList] Will insert after P${lastEditedIdx + 1}`);

          // Build all new paragraphs into a single OOXML package
          const newParagraphsXml = [];

          for (let i = existingCount; i < newCount; i++) {
            const item = itemsWithLevels[i];
            const ilvl = existingBaseIlvl + item.level;
            const numIdForPhase2 = (!existingNumId && listType === "numbered" && numFmt !== 'decimal') ? '100' : (existingNumId || templateNumId);

            console.log(`[executeEditList] Building paragraph ${i + 1}: "${item.text.substring(0, 30)}..." at ilvl=${ilvl}`);

            const escapedText = item.text
              .replace(/&/g, '&amp;')
              .replace(/</g, '&lt;')
              .replace(/>/g, '&gt;');

            // Build paragraph XML (without full package wrapper)
            newParagraphsXml.push(`
                      <w:p>
                        <w:pPr>
                          <w:pStyle w:val="ListParagraph"/>
                          <w:numPr>
                            <w:ilvl w:val="${ilvl}"/>
                            <w:numId w:val="${numIdForPhase2}"/>
                          </w:numPr>
                        </w:pPr>
                        <w:r>
                          <w:t xml:space="preserve">${escapedText}</w:t>
                        </w:r>
                      </w:p>`);
          }

          // Include numbering.xml part if using custom style
          let phase2NumberingPart = '';
          if (!existingNumId && listType === "numbered" && numFmt !== 'decimal') {
            phase2NumberingPart = `
          <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
            <pkg:xmlData>
              <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
            </pkg:xmlData>
          </pkg:part>
          <pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">
            <pkg:xmlData>
              <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:abstractNum w:abstractNumId="100">
                  <w:multiLevelType w:val="multilevel"/>
                  <w:lvl w:ilvl="0">
                    <w:start w:val="1"/>
                    <w:numFmt w:val="${numFmt}"/>
                    <w:lvlText w:val="%1."/>
                    <w:lvlJc w:val="left"/>
                    <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
                  </w:lvl>
                </w:abstractNum>
                <w:num w:numId="100">
                  <w:abstractNumId w:val="100"/>
                </w:num>
              </w:numbering>
            </pkg:xmlData>
          </pkg:part>`;
          }

          // Combine all paragraphs into a single OOXML package
          const combinedOxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
            <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
              <pkg:xmlData>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
              </pkg:xmlData>
            </pkg:part>${phase2NumberingPart}
            <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
              <pkg:xmlData>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    ${newParagraphsXml.join('\n')}
                  </w:body>
                </w:document>
              </pkg:xmlData>
            </pkg:part>
          </pkg:package>`;

          // Insert all new paragraphs at once
          const insertRange = insertAfterPara.getRange("End");
          insertRange.insertOoxml(combinedOxml, "After");
          await context.sync();

          console.log(`[executeEditList] Inserted ${newCount - existingCount} new paragraphs via single OOXML package`);

          // WORKAROUND: Insert a dummy spacing paragraph after the list, then remove it
          // This forces Word to properly re-evaluate and link the list structure
          try {
            // Reload paragraphs to get the newly inserted ones
            paragraphs.load("items/text");
            await context.sync();

            // Insert a blank paragraph after the last list item
            const lastListItemIdx = startIdx + newCount - 1;
            if (lastListItemIdx < paragraphs.items.length) {
              const lastListItem = paragraphs.items[lastListItemIdx];
              const dummyPara = lastListItem.insertParagraph("", "After");
              await context.sync();

              console.log(`[executeEditList] Inserted dummy spacing paragraph`);

              // Force Word to re-evaluate by syncing again
              await context.sync();

              // TEMP: Leave dummy paragraph to test if it fixes formatting
              // dummyPara.delete();
              // await context.sync();

              console.log(`[executeEditList] Left dummy spacing paragraph for testing`);
            }
          } catch (spacingError) {
            console.warn(`[executeEditList] Spacing workaround failed (non-critical):`, spacingError.message);
          }
        }

        // PHASE 3: Delete excess paragraphs if fewer new items
        if (newCount < existingCount) {
          // Delete from the end to avoid index shifting
          for (let i = existingCount - 1; i >= newCount; i--) {
            const paraToDelete = paragraphs.items[startIdx + i];
            paraToDelete.delete();
            console.log(`[executeEditList] Deleted excess P${startIdx + i + 1}`);
          }
          await context.sync();
        }

      } finally {
        await restoreChangeTracking(context, trackingState, "executeEditList");
      }

      console.log(`\n[executeEditList] ✅ SUCCESSFULLY COMPLETED`);
      console.log(`========== END EXECUTE_EDIT_LIST ==========\n\n`);
      console.log(`Successfully replaced paragraphs with ${listType} list`);
    });

    return {
      success: true,
      message: `Successfully created ${listType} list with ${newItems.length} items.`
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

  // Default to arabic if not specified
  const format = numberingFormat || "arabic";
  console.log(`executeConvertHeadersToList: Converting ${paragraphIndices.length} headers to ${format} numbered list`);

  try {
    await Word.run(async (context) => {
      const redlineEnabled = loadRedlineSetting();
      const trackingState = await setChangeTrackingForAi(context, redlineEnabled, "executeConvertHeadersToList");
      try {

        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items/text");
        await context.sync();

        // Sort indices to process in order
        const sortedIndices = [...paragraphIndices].sort((a, b) => a - b);

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
        let firstText = firstPara.text || "";
        const numberPattern = /^\s*\d+[.\)]\s*/;
        firstText = firstText.replace(numberPattern, "").trim();

        // Use new text if provided
        if (newHeaderTexts && newHeaderTexts.length > 0) {
          firstText = newHeaderTexts[0];
        }

        // Clear and replace the paragraph content
        firstPara.clear();
        firstPara.insertText(firstText, Word.InsertLocation.start);
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

          // Use new text if provided
          if (newHeaderTexts && newHeaderTexts.length > i) {
            paraText = newHeaderTexts[i];
          }

          // Clear and replace the paragraph content
          para.clear();
          para.insertText(paraText, Word.InsertLocation.start);
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

      console.log(`Successfully converted ${sortedIndices.length} headers to numbered list`);
    });

    return {
      success: true,
      message: `Successfully converted ${paragraphIndices.length} headers to a numbered list.`
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
      try {
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
        // Word Online requires items/ for rows
        table.load("rowCount, rows/items");
        await context.sync();

        if (action === "replace_content") {
          if (!content || !Array.isArray(content)) {
            throw new Error("replace_content requires a 2D array of content");
          }

          for (let r = 0; r < content.length && r < table.rows.items.length; r++) {
            table.rows.items[r].cells.load("items/body");
          }
          await context.sync();

          for (let r = 0; r < content.length && r < table.rows.items.length; r++) {
            const row = table.rows.items[r];
            for (let c = 0; c < content[r].length && c < row.cells.items.length; c++) {
              row.cells.items[c].load("body");
            }
          }
          await context.sync();

          for (let r = 0; r < content.length && r < table.rows.items.length; r++) {
            const row = table.rows.items[r];
            for (let c = 0; c < content[r].length && c < row.cells.items.length; c++) {
              const cell = row.cells.items[c];
              cell.body.clear();
              cell.body.insertText(content[r][c], Word.InsertLocation.start);
            }
          }
          await context.sync();

        } else if (action === "add_row") {
          if (!content || !Array.isArray(content)) {
            throw new Error("add_row requires an array of cell values");
          }
          const insertAt = targetRow !== undefined ? targetRow : table.rowCount;
          table.addRows(Word.InsertLocation.end, 1, [content]);
          await context.sync();

        } else if (action === "delete_row") {
          if (targetRow === undefined || targetRow < 0 || targetRow >= table.rows.items.length) {
            throw new Error(`Invalid or missing row index: ${targetRow}`);
          }
          table.rows.items[targetRow].delete();
          await context.sync();

        } else if (action === "update_cell") {
          if (targetRow === undefined || targetColumn === undefined) {
            throw new Error("update_cell requires targetRow and targetColumn");
          }
          if (targetRow < 0 || targetRow >= table.rows.items.length) {
            throw new Error(`Invalid row index: ${targetRow}`);
          }

          const row = table.rows.items[targetRow];
          row.cells.load("items/body");
          await context.sync();

          if (targetColumn < 0 || targetColumn >= row.cells.items.length) {
            throw new Error(`Invalid column index: ${targetColumn}`);
          }

          const cell = row.cells.items[targetColumn];
          cell.body.clear();
          cell.body.insertText(content[0], Word.InsertLocation.start);
          await context.sync();

        } else {
          throw new Error(`Unknown table action: ${action}`);
        }
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
