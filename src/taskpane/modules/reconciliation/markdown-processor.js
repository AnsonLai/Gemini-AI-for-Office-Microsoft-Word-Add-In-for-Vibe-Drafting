/**
 * OOXML Reconciliation Pipeline - Markdown Processor
 * 
 * Strips markdown syntax and captures format hints for later application.
 */

/**
 * Preprocesses markdown text by stripping formatting markers
 * and capturing their positions as format hints.
 * 
 * Supported formats:
 * - **bold** or __bold__
 * - *italic* or _italic_
 * - ++underline++
 * - ~~strikethrough~~
 * - ***bold+italic***
 * 
 * @param {string} text - Text with markdown formatting
 * @returns {import('./types.js').PreprocessResult}
 */
export function preprocessMarkdown(text) {
    if (!text) {
        return { cleanText: '', formatHints: [] };
    }

    const formatHints = [];
    let cleanText = '';
    let lastIndex = 0;

    // Process patterns in order of specificity (longer patterns first)
    // Pattern groups capture: full match, inner text
    const patterns = [
        // HTML Bold: <b>text</b>, <strong>text</strong>
        { regex: /<b>(.+?)<\/b>/gi, format: { bold: true } },
        { regex: /<strong>(.+?)<\/strong>/gi, format: { bold: true } },
        // HTML Italic: <i>text</i>, <em>text</em>
        { regex: /<i>(.+?)<\/i>/gi, format: { italic: true } },
        { regex: /<em>(.+?)<\/em>/gi, format: { italic: true } },
        // HTML Underline: <u>text</u>
        { regex: /<u>(.+?)<\/u>/gi, format: { underline: true } },
        // HTML Strikethrough: <s>text</s>, <strike>text</strike>, <del>text</del>
        { regex: /<s>(.+?)<\/s>/gi, format: { strikethrough: true } },
        { regex: /<strike>(.+?)<\/strike>/gi, format: { strikethrough: true } },
        { regex: /<del>(.+?)<\/del>/gi, format: { strikethrough: true } },

        // Escaped HTML Bold: &lt;b&gt;text&lt;/b&gt;, &lt;strong&gt;text&lt;/strong&gt;
        { regex: /&lt;b&gt;(.+?)&lt;\/b&gt;/gi, format: { bold: true }, isEscaped: true },
        { regex: /&lt;strong&gt;(.+?)&lt;\/strong&gt;/gi, format: { bold: true }, isEscaped: true },
        // Escaped HTML Italic: &lt;i&gt;text&lt;/i&gt;, &lt;em&gt;text&lt;/em&gt;
        { regex: /&lt;i&gt;(.+?)&lt;\/i&gt;/gi, format: { italic: true }, isEscaped: true },
        { regex: /&lt;em&gt;(.+?)&lt;\/em&gt;/gi, format: { italic: true }, isEscaped: true },
        // Escaped HTML Underline: &lt;u&gt;text&lt;/u&gt;
        { regex: /&lt;u&gt;(.+?)&lt;\/u&gt;/gi, format: { underline: true }, isEscaped: true },
        // Escaped HTML Strikethrough: &lt;s&gt;text&lt;/s&gt;
        { regex: /&lt;s&gt;(.+?)&lt;\/s&gt;/gi, format: { strikethrough: true }, isEscaped: true },

        // Bold + Italic: ***text***
        { regex: /\*\*\*(.+?)\*\*\*/g, format: { bold: true, italic: true } },
        // Bold + Underline: **++text++**
        { regex: /\*\*\+\+(.+?)\+\+\*\*/g, format: { bold: true, underline: true } },
        // Bold: **text** or __text__
        { regex: /\*\*(.+?)\*\*/g, format: { bold: true } },
        { regex: /__(.+?)__/g, format: { bold: true } },
        // Underline: ++text++
        { regex: /\+\+(.+?)\+\+/g, format: { underline: true } },
        // Strikethrough: ~~text~~
        { regex: /~~(.+?)~~/g, format: { strikethrough: true } },
        // Italic: *text* or _text_ (must come after ** and __)
        { regex: /(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, format: { italic: true } },
        { regex: /(?<!_)_(?!_)(.+?)(?<!_)_(?!_)/g, format: { italic: true } }
    ];

    // Build a combined approach: find all matches, then process in order
    const allMatches = [];

    for (const pattern of patterns) {
        let match;
        const regex = new RegExp(pattern.regex.source, pattern.regex.flags);
        while ((match = regex.exec(text)) !== null) {
            allMatches.push({
                start: match.index,
                end: match.index + match[0].length,
                fullMatch: match[0],
                innerText: pattern.isEscaped ? decodeHtmlEntities(match[1]) : match[1],
                format: pattern.format
            });
        }
    }

    // Sort by start position
    allMatches.sort((a, b) => a.start - b.start);

    // Remove overlapping matches (keep the first one)
    const filteredMatches = [];
    let lastEnd = 0;
    for (const match of allMatches) {
        if (match.start >= lastEnd) {
            filteredMatches.push(match);
            lastEnd = match.end;
        }
    }

    // Build clean text and format hints
    let offset = 0;
    lastIndex = 0;

    for (const match of filteredMatches) {
        // Add text before this match
        const beforeText = text.slice(lastIndex, match.start);
        cleanText += beforeText;
        offset += beforeText.length;

        // Add the inner text (without markers)
        const innerStart = offset;
        cleanText += match.innerText;
        const innerEnd = offset + match.innerText.length;
        offset = innerEnd;

        // Record the format hint
        formatHints.push({
            start: innerStart,
            end: innerEnd,
            format: match.format
        });

        lastIndex = match.end;
    }

    // Add remaining text after last match
    cleanText += text.slice(lastIndex);

    return { cleanText, formatHints };
}

/**
 * Checks if any format hints apply to a given offset range.
 * 
 * @param {import('./types.js').FormatHint[]} formatHints - Array of format hints
 * @param {number} startOffset - Start of range to check
 * @param {number} endOffset - End of range to check
 * @returns {import('./types.js').FormatHint[]} Applicable hints
 */
export function getApplicableFormatHints(formatHints, startOffset, endOffset) {
    return formatHints.filter(hint =>
        hint.start < endOffset && hint.end > startOffset
    );
}

/**
 * Merges format objects (combines multiple format flags).
 * 
 * @param  {...Object} formats - Format objects to merge
 * @returns {Object} Combined format object
 */
export function mergeFormats(...formats) {
    const result = {};
    for (const format of formats) {
        if (format) {
            Object.assign(result, format);
        }
    }
    return result;
}

/**
 * Decodes HTML entities in text (e.g. &amp; -> &)
 */
function decodeHtmlEntities(text) {
    if (!text) return '';
    return text
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#039;/g, "'");
}
