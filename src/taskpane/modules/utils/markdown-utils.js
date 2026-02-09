/* global Word */

import { marked } from 'marked';
import { preprocessMarkdown, parseMarkdownListContent } from '../reconciliation/index.js';

// Cached document font - set by detectDocumentFont() before edits
let cachedDocumentFont = "Calibri"; // Safe default for Word

/**
 * Detects and caches the document's font from the first paragraph.
 * Should be called before making edits to ensure font consistency.
 */
async function detectDocumentFont() {
  try {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/font/name");
      await context.sync();

      if (paragraphs.items.length > 0) {
        const firstPara = paragraphs.items[0];
        firstPara.load("font/name");
        await context.sync();

        if (firstPara.font.name) {
          cachedDocumentFont = firstPara.font.name;
          console.log(`Detected document font: ${cachedDocumentFont}`);
        }
      }
    });
  } catch (error) {
    console.warn("Could not detect document font, using default:", error);
  }
  return cachedDocumentFont;
}

/**
 * Converts markdown to Word-compatible HTML.
 * Ensures proper formatting for Word's HTML parser.
 */
function markdownToWordHtml(markdown) {
  console.log(`[markdownToWordHtml] Processing content: ${markdown.substring(0, 50)}...`);
  if (!markdown) return "";

  // Pre-process underline and strikethrough (marked native GFM might be disabled or fail in some environments)
  const processedMarkdown = markdown
    .replace(/^\s*(#{7,9})\s+(.*)$/gm, (match, hashes, text) => {
      const level = Math.min(hashes.length, 9);
      return `<p style="mso-style-name:'Heading ${level}'; mso-style-id:Heading${level}; font-weight:bold;">${escapeHtml(text.trim())}</p>`;
    })
    .replace(/\+\+(.+?)\+\+/g, '<u>$1</u>')
    .replace(/~~(.+?)~~/g, '<s>$1</s>');

  // Parse markdown to HTML using marked library
  let html = marked.parse(processedMarkdown);

  // === TABLE FORMATTING ===
  // Word requires explicit styling for tables to render properly with borders
  if (html.includes('<table>')) {
    html = html.replace(/<table>/g, '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border: 1px solid #000;">');
    // Add styling to table cells for better appearance
    html = html.replace(/<th>/g, '<th style="border: 1px solid #000; padding: 8px; background-color: #f0f0f0; font-weight: bold;">');
    html = html.replace(/<td>/g, '<td style="border: 1px solid #000; padding: 8px;">');
  }

  // === ORDERED LIST FORMATTING ===
  // CRITICAL: Word's HTML parser can render <ol> as bullets without explicit styling
  // Adding list-style-type CSS ensures proper numbered list rendering
  if (html.includes('<ol>')) {
    // First, handle any already-styled ordered lists (from nested replacements)
    // to avoid double-replacing them
    html = html.replace(/<ol>/g, '<ol style="list-style-type: decimal; margin-left: 0; padding-left: 40px; margin-bottom: 10px;">');

    // Handle nested ordered lists with different numbering styles
    // Match <ol> tags that are inside <li> elements (nested lists)
    // Use lower-alpha (a, b, c) for first nesting level
    html = html.replace(/<li>([^<]*)<ol style="list-style-type: decimal;/g, '<li>$1<ol style="list-style-type: lower-alpha;');

    // For third-level nesting, use lower-roman (i, ii, iii)
    html = html.replace(/<li>([^<]*)<ol style="list-style-type: lower-alpha;([^>]*)>([^<]*)<li>([^<]*)<ol style="list-style-type: lower-alpha;/g,
      '<li>$1<ol style="list-style-type: lower-alpha;$2>$3<li>$4<ol style="list-style-type: lower-roman;');
  }

  // === UNORDERED LIST FORMATTING ===
  // Ensure <ul> has explicit bullet styling to distinguish from ordered lists
  if (html.includes('<ul>')) {
    html = html.replace(/<ul>/g, '<ul style="list-style-type: disc; margin-left: 0; padding-left: 40px; margin-bottom: 10px;">');

    // Nested unordered lists should use circle then square markers
    html = html.replace(/<li>([^<]*)<ul style="list-style-type: disc;/g, '<li>$1<ul style="list-style-type: circle;');
    html = html.replace(/<li>([^<]*)<ul style="list-style-type: circle;/g, '<li>$1<ul style="list-style-type: square;');
  }

  // === LIST ITEM FORMATTING ===
  // Add spacing to list items for better readability
  html = html.replace(/<li>/g, '<li style="margin-bottom: 5px;">');

  // === TRAILING PARAGRAPH ===
  // Add a paragraph with non-breaking space after lists to ensure proper formatting
  // This fixes the issue where the last list item may not be properly numbered/formatted
  // Using &nbsp; instead of empty paragraph for better Word compatibility
  html = html.replace(/<\/ol>/g, '</ol><p>&nbsp;</p>');
  html = html.replace(/<\/ul>/g, '</ul><p>&nbsp;</p>');

  // === FONT CONSISTENCY ===
  // Use a block wrapper when content includes block elements to avoid invalid HTML
  const hasBlockHtml = /<(p|ul|ol|table|h[1-6]|div|blockquote)\b/i.test(html);
  const wrapperTag = hasBlockHtml ? 'div' : 'span';
  html = `<${wrapperTag} style="font-family: '${cachedDocumentFont}', Calibri, sans-serif;">${html}</${wrapperTag}>`;

  return html;
}

/**
 * Converts markdown to Word-compatible HTML for inline content (no wrapping <p> tags).
 * Use this for modify_text replacements.
 */
function markdownToWordHtmlInline(markdown) {
  if (!markdown) return "";

  // Use parseInline to avoid wrapping in <p> tags for simple text
  // But if there are block elements (lists, tables), use full parse
  const hasBlockElements = /(^|\n)\s*(?:[-*+•]|\d+(?:\.\d+)*\.?|[A-Za-z]\.|[ivxlcIVXLC]+\.)\s+|(\|.*\|.*\n)|(^#{1,9}\s)/m.test(markdown);

  if (hasBlockElements) {
    return markdownToWordHtml(markdown);
  }

  // For inline content, pre-process underline/strike and use parseInline
  const processedMarkdown = markdown
    .replace(/\+\+(.+?)\+\+/g, '<u>$1</u>')
    .replace(/~~(.+?)~~/g, '<s>$1</s>');
  return `<span style="font-family: '${cachedDocumentFont}', Calibri, sans-serif;">${marked.parseInline(processedMarkdown)}</span>`;
}

/**
 * Detects if content has block elements (lists, tables, headings)
 * that require HTML insertion instead of word-level diffs
 */
function hasBlockElements(content) {
  if (!content) return false;

  // Check for markdown block elements with improved patterns

  // Detect unordered lists: lines starting with -, *, or + followed by space
  const hasUnorderedList = /^[\s]*[-*+]\s+/m.test(content);

  // Detect ordered lists: lines starting with number(s) followed by period and space
  // Examples: "1. item", "10. item", "  2. item"
  const hasOrderedList = /^[\s]*\d+\.\s+/m.test(content);

  // Detect outline numbering and alpha/roman markers
  // Examples: "1.1. item", "A. item", "a. item", "I. item", "iv. item"
  const hasOutlineList = /^[\s]*\d+\.\d+(?:\.\d+)*\.?\s+/m.test(content);
  const hasAlphaDotList = /^[\s]*[A-Za-z]\.\s+/m.test(content);
  const hasRomanDotList = /^[\s]*[ivxlcIVXLC]+\.\s+/m.test(content);

  // Detect alphabetical lists: (a), (b), (c) style
  const hasAlphaList = /^[\s]*\([a-z]\)\s+/m.test(content);

  // Detect tables: markdown table syntax with pipes
  const hasTable = /\|.*\|.*\n/.test(content);

  // Detect headings: lines starting with # symbols
  const hasHeading = /^#{1,9}\s/m.test(content);

  // Detect paragraph breaks (multiple consecutive newlines)
  const hasMultipleLineBreaks = content.includes('\n\n');

  const result = hasUnorderedList || hasOrderedList || hasOutlineList || hasAlphaDotList || hasRomanDotList || hasAlphaList || hasTable || hasHeading || hasMultipleLineBreaks;

  // Debug logging to help diagnose issues
  if (result) {
    console.log('Block elements detected:', {
      hasUnorderedList,
      hasOrderedList,
      hasOutlineList,
      hasAlphaDotList,
      hasRomanDotList,
      hasAlphaList,
      hasTable,
      hasHeading,
      hasMultipleLineBreaks,
      contentPreview: content.substring(0, 100)
    });
  }

  return result;
}

/**
 * Checks if text contains inline markdown formatting (bold, italic, code, etc.)
 * Returns true if formatting patterns are detected
 */
function hasInlineMarkdownFormatting(text) {
  if (!text) return false;
  // Check for common inline markdown patterns:
  // **bold**, *italic*, __bold__, _italic_, `code`, ~~strikethrough~~, ++underline++
  // Also check for **...** pattern specifically
  return /(\*\*.+?\*\*|\*.+?\*|__.+?__|_.+?_|`.+?`|~~.+?~~|\+\+.+?\+\+)/.test(text);
}

/**
 * Wrapper for preprocessMarkdown that handles empty paragraph cases.
 * Returns clean text and format hints for formatting application.
 * 
 * @param {string} content - Content with markdown formatting
 * @returns {{ cleanText: string, formatHints: Array }}
 */
async function preprocessMarkdownForParagraph(content) {
  try {
    return preprocessMarkdown(content);
  } catch (e) {
    console.error('preprocessMarkdown failed:', e);
    return { cleanText: content, formatHints: [] };
  }
}

/**
 * Applies format hints to specific text ranges in a paragraph using Word's font API.
 * This avoids HTML/OOXML insertion issues in table cells.
 * 
 * @param {Word.Paragraph} paragraph - Target paragraph
 * @param {string} text - The text content
 * @param {Array} formatHints - Array of format hints with start/end/format
 * @param {Word.RequestContext} context - Word context
 */
async function applyFormatHintsToRanges(paragraph, text, formatHints, context) {
  // Load paragraph as range
  const paragraphRange = paragraph.getRange();
  paragraphRange.load('text');
  await context.sync();

  // Get the paragraph text to verify positions
  const paragraphText = paragraphRange.text;

  for (const hint of formatHints) {
    try {
      // Calculate the text to search for based on the hint offsets
      const hintText = text.substring(hint.start, hint.end);
      if (!hintText || hintText.trim().length === 0) continue;

      // Search for the text within the paragraph
      const searchResults = paragraphRange.search(hintText, { matchCase: true, matchWholeWord: false });
      searchResults.load('items/text');
      await context.sync();

      if (searchResults.items.length > 0) {
        // Apply formatting to the first match
        const targetRange = searchResults.items[0];

        if (hint.format.bold) {
          targetRange.font.bold = true;
        }
        if (hint.format.italic) {
          targetRange.font.italic = true;
        }
        if (hint.format.underline) {
          targetRange.font.underline = Word.UnderlineType.single;
        }
        if (hint.format.strikethrough) {
          targetRange.font.strikeThrough = true;
        }

        await context.sync();
      }
    } catch (formatError) {
      console.warn(`Could not apply formatting to hint at ${hint.start}-${hint.end}:`, formatError);
    }
  }
}

/**
 * Removes formatting from specific text ranges using Word's native Font API.
 * This is used when the AI sends plain text to remove formatting (e.g., unbold).
 * Using the native Font API allows Word to properly track format changes.
 * 
 * @param {Word.Paragraph} paragraph - Target paragraph
 * @param {string} text - The text content
 * @param {Array} formatRemovalHints - Array of hints with start/end/removeFormat
 * @param {Word.RequestContext} context - Word context
 */
async function applyFormatRemovalToRanges(paragraph, text, formatRemovalHints, context) {
  // Load paragraph as range
  const paragraphRange = paragraph.getRange();
  paragraphRange.load('text');
  await context.sync();

  console.log(`[FontAPI] Applying format removal to ${formatRemovalHints.length} ranges`);

  for (const hint of formatRemovalHints) {
    try {
      // Calculate the text to search for based on the hint offsets
      const hintText = text.substring(hint.start, hint.end);
      if (!hintText || hintText.trim().length === 0) continue;

      console.log(`[FontAPI] Searching for "${hintText}" to remove format:`, hint.removeFormat);

      // Search for the text within the paragraph
      const searchResults = paragraphRange.search(hintText, { matchCase: true, matchWholeWord: false });
      searchResults.load('items/text');
      await context.sync();

      if (searchResults.items.length > 0) {
        // Apply format removal to the first match
        const targetRange = searchResults.items[0];

        // Remove formatting by setting to false
        if (hint.removeFormat.bold) {
          targetRange.font.bold = false;
          console.log(`[FontAPI] Set bold=false for "${hintText}"`);
        }
        if (hint.removeFormat.italic) {
          targetRange.font.italic = false;
          console.log(`[FontAPI] Set italic=false for "${hintText}"`);
        }
        if (hint.removeFormat.underline) {
          targetRange.font.underline = Word.UnderlineType.none;
          console.log(`[FontAPI] Set underline=none for "${hintText}"`);
        }
        if (hint.removeFormat.strikethrough) {
          targetRange.font.strikeThrough = false;
          console.log(`[FontAPI] Set strikeThrough=false for "${hintText}"`);
        }

        await context.sync();
        console.log(`[FontAPI] Successfully removed formatting from "${hintText}"`);
      } else {
        console.warn(`[FontAPI] Text not found: "${hintText}"`);
      }
    } catch (formatError) {
      console.warn(`Could not remove formatting at ${hint.start}-${hint.end}:`, formatError);
    }
  }
}

/**
 * Parses markdown list content into structured data
 * Supports numbered lists (1. item) and bullet lists (- item, * item)
 */
function parseMarkdownList(content) {
  const parsed = parseMarkdownListContent(content);
  if (!parsed) return null;

  const lines = String(content || '').trim().split('\n').filter(line => line.trim().length > 0);
  const hasNumbered = parsed.items.some(i => i.type === 'numbered');
  const hasBullet = parsed.items.some(i => i.type === 'bullet');
  console.log(`[parseMarkdownList] Processing ${lines.length} lines. hasNumbered=${hasNumbered}, hasBullet=${hasBullet}`);

  return parsed;
}

/**
 * Normalizes content by converting literal escape sequences to actual characters.
 * This is necessary because AI responses sometimes return "\\n" as a two-character
 * string instead of actual newlines, which breaks markdown parsing.
 */
function normalizeContentEscapes(content) {
  if (!content || typeof content !== 'string') return content;

  // Convert literal \n (two characters) to actual newline
  // Also handle other common escapes
  return content
    .replace(/\\n/g, '\n')      // Literal \n -> newline
    .replace(/\\t/g, '\t')      // Literal \t -> tab
    .replace(/\\r/g, '\r');     // Literal \r -> carriage return
}

export {
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
};

function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
