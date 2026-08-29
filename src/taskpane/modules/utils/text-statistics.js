/**
 * text-statistics.js
 * Utility functions for stripping formatting and calculating document / selection statistics
 * (word count, character count with and without spaces, paragraphs, sentences, reading time).
 */

/**
 * Strips all formatting, Markdown syntax, internal document annotations,
 * HTML/XML tags, and Unicode anomalies from raw text.
 *
 * @param {string} rawText - Text containing potential formatting/markers
 * @returns {string} - Clean, plain text
 */
export function stripFormatting(rawText) {
  if (typeof rawText !== 'string' || !rawText) {
    return '';
  }

  let text = rawText;

  // 1. Remove internal plugin paragraph/metadata anchors: [P1|Normal], [P2|ListNumber|L:0|§], [T:row,col], [P3]
  text = text.replace(/\[P\d+[^\]]*\]/g, '');

  // 2. Remove HTML/XML tags (<span...>, <p>, </w:t>, etc.) before unescaping entities
  text = text.replace(/<[^>]+>/g, ' ');

  // 3. Decode common HTML entities
  text = text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&apos;/gi, "'");

  // 4. Normalize newlines and special whitespace
  text = text
    .replace(/\r\n|\r|\v|\f/g, '\n')
    // Remove zero-width spaces, joiners, byte-order marks, soft hyphens
    .replace(/[\u200B-\u200D\uFEFF\u00AD]/g, '')
    // Replace non-breaking spaces with standard space
    .replace(/[\u00A0\u202F]/g, ' ');

  // 5. Remove Markdown Code Blocks (```lang\n...\n```)
  text = text.replace(/```[\s\S]*?```/g, (match) => {
    // Keep internal text of code block without backticks/language identifier
    return match.replace(/^```[^\n]*\n?/gm, '').replace(/```$/gm, '');
  });

  // 6. Remove Markdown Inline Code (`code`)
  text = text.replace(/`([^`\n]+)`/g, '$1');

  // 7. Remove Markdown Links and Images: [text](url) -> text, ![alt](url) -> alt
  text = text.replace(/!\[([^\]]*)\]\([^)]*\)/g, '$1');
  text = text.replace(/\[([^\]]+)\]\([^)]*\)/g, '$1');

  // 8. Remove Markdown Headings (# Heading)
  text = text.replace(/^\s*#{1,6}\s+(.*)$/gm, '$1');

  // 9. Remove Markdown Blockquotes (> quote)
  text = text.replace(/^\s*>\s+(.*)$/gm, '$1');

  // 10. Remove Markdown Horizontal Rules (---, ***, ___)
  text = text.replace(/^\s*[-*_]{3,}\s*$/gm, '');

  // 11. Remove Markdown Table Borders and Separator Lines (|---|---|)
  text = text.replace(/^\s*\|?(?:\s*:?-+:?\s*\|)+\s*$/gm, ''); // Table separator rows
  text = text.replace(/^\s*\|/gm, '').replace(/\|\s*$/gm, '');  // Leading/trailing table pipes
  text = text.replace(/\|/g, ' ');                              // Column separators to spaces

  // 12. Remove Markdown List Markers (*, -, +, 1., a., i., 1.1.)
  text = text.replace(/^\s*(?:[-*+]|\d+[\.\)]|[a-zA-Z][\.\)]|\d+(?:\.\d+)+[\.\)]?)\s+/gm, '');

  // 13. Remove Inline Formatting Tokens while preserving inner text
  // Bold & Italic: ***text*** or ___text___
  text = text.replace(/(\*{3}|_{3})(.*?)\1/g, '$2');
  // Bold: **text** or __text__
  text = text.replace(/(\*{2}|_{2})(.*?)\1/g, '$2');
  // Italic: *text* or _text_
  text = text.replace(/(\*|_)(.*?)\1/g, '$2');
  // Underline: ++text++
  text = text.replace(/\+\+(.*?)\+\+/g, '$1');
  // Strikethrough: ~~text~~
  text = text.replace(/~~(.*?)~~/g, '$1');

  // 14. Clean up extraneous whitespace within lines
  text = text
    .split('\n')
    .map(line => line.replace(/[ \t]+/g, ' ').trim())
    .filter((line, idx, arr) => {
      // Collapse excessive blank lines
      if (!line && idx > 0 && !arr[idx - 1]) return false;
      return true;
    })
    .join('\n')
    .trim();

  return text;
}

/**
 * Counts the words in clean text using Intl.Segmenter with regex fallback.
 *
 * @param {string} text - Clean text
 * @returns {number} - Total word count
 */
export function countWords(text) {
  if (!text || !text.trim()) {
    return 0;
  }

  const trimmed = text.trim();

  if (typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function') {
    try {
      const segmenter = new Intl.Segmenter(undefined, { granularity: 'word' });
      let count = 0;
      for (const segment of segmenter.segment(trimmed)) {
        if (segment.isWordLike) {
          count++;
        }
      }
      return count;
    } catch (_) {
      // Fall through to regex
    }
  }

  // Regex fallback: matches unicode word tokens (letters, digits, accented chars)
  // including embedded hyphens and apostrophes (e.g. state-of-the-art, don't)
  const matches = trimmed.match(/[\p{L}\p{N}]+(?:['’\-][\p{L}\p{N}]+)*/gu);
  return matches ? matches.length : 0;
}

/**
 * Counts the sentences in clean text.
 *
 * @param {string} text - Clean text
 * @returns {number} - Total sentence count
 */
export function countSentences(text) {
  if (!text || !text.trim()) {
    return 0;
  }

  const trimmed = text.trim();

  if (typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function') {
    try {
      const segmenter = new Intl.Segmenter(undefined, { granularity: 'sentence' });
      let count = 0;
      for (const segment of segmenter.segment(trimmed)) {
        if (segment.segment.trim().length > 0) {
          count++;
        }
      }
      return count;
    } catch (_) {
      // Fall through to regex
    }
  }

  const matches = trimmed.match(/[^.!?]+(?:[.!?]+["'”’]?|$)/g);
  return matches ? matches.filter(s => s.trim().length > 0).length : 0;
}

/**
 * Calculates full text statistics for a given raw string.
 *
 * @param {string} rawText - Raw input text (selection or document content)
 * @returns {object} - Object containing word count, character count, etc.
 */
export function calculateTextStatistics(rawText) {
  const cleanText = stripFormatting(rawText);

  if (!cleanText) {
    return {
      wordCount: 0,
      characterCountWithSpaces: 0,
      characterCountWithoutSpaces: 0,
      paragraphCount: 0,
      sentenceCount: 0,
      estimatedReadingTime: "0 sec",
      estimatedReadingTimeMinutes: 0,
      preview: ""
    };
  }

  const wordCount = countWords(cleanText);
  const characterCountWithSpaces = cleanText.length;
  const characterCountWithoutSpaces = cleanText.replace(/\s/g, '').length;

  const paragraphs = cleanText
    .split(/\n+/)
    .map(p => p.trim())
    .filter(p => p.length > 0);
  const paragraphCount = paragraphs.length;

  const sentenceCount = countSentences(cleanText);

  // Estimated reading time (~200 words per minute)
  const totalReadingSeconds = Math.round((wordCount / 200) * 60);
  let estimatedReadingTime = `${totalReadingSeconds} sec`;
  if (totalReadingSeconds >= 60) {
    const mins = (totalReadingSeconds / 60).toFixed(1);
    estimatedReadingTime = `${mins} min (${totalReadingSeconds} sec)`;
  }

  const preview = cleanText.length > 120
    ? cleanText.substring(0, 117) + '...'
    : cleanText;

  return {
    wordCount,
    characterCountWithSpaces,
    characterCountWithoutSpaces,
    paragraphCount,
    sentenceCount,
    estimatedReadingTime,
    estimatedReadingTimeMinutes: parseFloat((wordCount / 200).toFixed(2)),
    preview
  };
}
