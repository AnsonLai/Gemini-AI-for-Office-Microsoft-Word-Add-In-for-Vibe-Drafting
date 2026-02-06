/**
 * OOXML Reconciliation Pipeline - Markdown Processor
 * 
 * Strips markdown syntax and captures format hints for later application.
 */

const MARKDOWN_PATTERNS = [
    // HTML Bold: <b>text</b>, <strong>text</strong>
    { regex: /<b>(.+?)<\/b>/i, format: { bold: true } },
    { regex: /<strong>(.+?)<\/strong>/i, format: { bold: true } },
    // HTML Italic: <i>text</i>, <em>text</em>
    { regex: /<i>(.+?)<\/i>/i, format: { italic: true } },
    { regex: /<em>(.+?)<\/em>/i, format: { italic: true } },
    // HTML Underline: <u>text</u>
    { regex: /<u>(.+?)<\/u>/i, format: { underline: true } },
    // HTML Strikethrough: <s>text</s>, <strike>text</strike>, <del>text</del>
    { regex: /<s>(.+?)<\/s>/i, format: { strikethrough: true } },
    { regex: /<strike>(.+?)<\/strike>/i, format: { strikethrough: true } },
    { regex: /<del>(.+?)<\/del>/i, format: { strikethrough: true } },

    // Escaped HTML Bold: &lt;b&gt;text&lt;/b&gt;, &lt;strong&gt;text&lt;/strong&gt;
    { regex: /&lt;b&gt;(.+?)&lt;\/b&gt;/i, format: { bold: true }, isEscaped: true },
    { regex: /&lt;strong&gt;(.+?)&lt;\/strong&gt;/i, format: { bold: true }, isEscaped: true },
    // Escaped HTML Italic: &lt;i&gt;text&lt;/i&gt;, &lt;em&gt;text&lt;/em&gt;
    { regex: /&lt;i&gt;(.+?)&lt;\/i&gt;/i, format: { italic: true }, isEscaped: true },
    { regex: /&lt;em&gt;(.+?)&lt;\/em&gt;/i, format: { italic: true }, isEscaped: true },
    // Escaped HTML Underline: &lt;u&gt;text&lt;/u&gt;
    { regex: /&lt;u&gt;(.+?)&lt;\/u&gt;/i, format: { underline: true }, isEscaped: true },
    // Escaped HTML Strikethrough: &lt;s&gt;text&lt;/s&gt;
    { regex: /&lt;s&gt;(.+?)&lt;\/s&gt;/i, format: { strikethrough: true }, isEscaped: true },

    // Bold + Italic: ***text***
    { regex: /\*\*\*(.+?)\*\*\*/, format: { bold: true, italic: true } },
    // Bold + Underline: **++text++**
    { regex: /\*\*\+\+(.+?)\+\+\*\*/, format: { bold: true, underline: true } },
    // Bold: **text** or __text__
    { regex: /\*\*(.+?)\*\*/, format: { bold: true } },
    { regex: /__(.+?)__/, format: { bold: true } },
    // Underline: ++text++
    { regex: /\+\+(.+?)\+\+/, format: { underline: true } },
    // Strikethrough: ~~text~~ or ~text~
    { regex: /~~(.+?)~~/, format: { strikethrough: true } },
    { regex: /~(.+?)~/, format: { strikethrough: true } },
    // Italic: *text* or _text_ (using lookahead only for compatibility)
    { regex: /\*(?!\*)(.+?)\*(?!\*)/, format: { italic: true } },
    { regex: /_(?!_)(.+?)_(?!_)/, format: { italic: true } }
];

/**
 * Preprocesses markdown text by stripping formatting markers
 * and capturing their positions as format hints.
 */
export function preprocessMarkdown(text) {
    if (!text) {
        return { cleanText: '', formatHints: [] };
    }

    const formatHints = [];
    let cleanText = '';

    // Find all matches for all patterns
    const allMatches = [];
    for (const pattern of MARKDOWN_PATTERNS) {
        let match;
        const source = pattern.regex.source || pattern.regex.toString().replace(/^\/|\/[gimuy]*$/g, '');
        const flags = 'g' + (pattern.regex.ignoreCase ? 'i' : '');
        const regex = new RegExp(source, flags);

        while ((match = regex.exec(text)) !== null) {
            allMatches.push({
                start: match.index,
                end: match.index + match[0].length,
                fullMatch: match[0],
                innerText: pattern.isEscaped ? decodeHtmlEntities(match[1]) : match[1],
                format: pattern.format
            });
            if (match.index === regex.lastIndex) regex.lastIndex++;
        }
    }

    // Sort: earliest first, then longest first
    allMatches.sort((a, b) => (a.start - b.start) || (b.end - a.end));

    // Filter to keep only top-level matches
    const topLevelMatches = [];
    let lastEnd = 0;
    for (const match of allMatches) {
        if (match.start >= lastEnd) {
            topLevelMatches.push(match);
            lastEnd = match.end;
        }
    }

    // Recursive reconstruction
    let lastIndex = 0;
    for (const match of topLevelMatches) {
        cleanText += text.slice(lastIndex, match.start);

        const subResult = preprocessMarkdown(match.innerText);

        const segmentStart = cleanText.length;
        cleanText += subResult.cleanText;
        const segmentEnd = cleanText.length;

        formatHints.push({
            start: segmentStart,
            end: segmentEnd,
            format: match.format
        });

        for (const subHint of subResult.formatHints) {
            formatHints.push({
                start: segmentStart + subHint.start,
                end: segmentStart + subHint.end,
                format: subHint.format
            });
        }

        lastIndex = match.end;
    }

    cleanText += text.slice(lastIndex);
    return { cleanText, formatHints };
}

/**
 * Checks if any format hints apply to a given offset range.
 */
export function getApplicableFormatHints(formatHints, startOffset, endOffset) {
    return formatHints.filter(hint =>
        hint.start < endOffset && hint.end > startOffset
    );
}

/**
 * Merges format objects.
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
 * Decodes HTML entities in text.
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

