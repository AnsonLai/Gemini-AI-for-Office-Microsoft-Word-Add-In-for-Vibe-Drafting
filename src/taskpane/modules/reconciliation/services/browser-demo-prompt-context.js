/**
 * Builds prompt sections for browser-demo paragraph context.
 *
 * Plain lines are always emitted as `[P#] ...`.
 * Formatting lines are emitted as `[P#_FMT] ...` only when markdown projection
 * differs from plain text.
 *
 * @param {Array<{ index?: number, text?: string, formattedText?: string }>} paragraphs
 * @returns {{ plainListing: string, formattingListing: string }}
 */
export function buildPromptParagraphSections(paragraphs) {
    const items = Array.isArray(paragraphs) ? paragraphs : [];
    const plainLines = [];
    const formattingLines = [];

    for (let i = 0; i < items.length; i += 1) {
        const paragraph = items[i] || {};
        const index = Number.isInteger(paragraph.index) && paragraph.index > 0 ? paragraph.index : (i + 1);
        const text = String(paragraph.text || '').trim();
        if (!text) continue;

        const formattedText = String(paragraph.formattedText || text).trim() || text;
        plainLines.push(`[P${index}] ${text}`);
        if (formattedText !== text) {
            formattingLines.push(`[P${index}_FMT] ${formattedText}`);
        }
    }

    return {
        plainListing: plainLines.join('\n'),
        formattingListing: formattingLines.join('\n')
    };
}

function normalizeWhitespace(value) {
    return String(value || '').replace(/\s+/g, ' ').trim();
}

function truncateForLog(value, maxChars = 140) {
    const text = normalizeWhitespace(value);
    if (text.length <= maxChars) return text;
    return `${text.slice(0, maxChars - 1)}…`;
}

function extractQueryPhrases(userMessage) {
    const message = String(userMessage || '');
    const queries = [];
    const seen = new Set();
    const addQuery = (raw) => {
        const normalized = normalizeWhitespace(raw);
        if (!normalized || normalized.length < 2) return;
        const key = normalized.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        queries.push(normalized);
    };

    const quotedRegex = /"([^"]+)"|'([^']+)'/g;
    let quotedMatch = quotedRegex.exec(message);
    while (quotedMatch) {
        addQuery(quotedMatch[1] || quotedMatch[2]);
        quotedMatch = quotedRegex.exec(message);
    }

    const capitalizedPhraseRegex = /\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\b/g;
    let capMatch = capitalizedPhraseRegex.exec(message);
    while (capMatch) {
        addQuery(capMatch[1]);
        capMatch = capitalizedPhraseRegex.exec(message);
    }

    return queries.slice(0, 8);
}

function extractQuotedPhrases(text) {
    const message = String(text || '');
    const phrases = [];
    const seen = new Set();
    const add = (raw) => {
        const normalized = normalizeWhitespace(raw);
        if (!normalized || normalized.length < 2) return;
        const key = normalized.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        phrases.push(normalized);
    };

    const regexes = [
        /"([^"]+)"/g,
        /'([^']+)'/g,
        /“([^”]+)”/g,
        /‘([^’]+)’/g
    ];
    for (const pattern of regexes) {
        let match = pattern.exec(message);
        while (match) {
            add(match[1]);
            match = pattern.exec(message);
        }
    }

    return phrases;
}

const MESSAGE_STOP_WORDS = new Set([
    'a', 'an', 'and', 'as', 'at', 'be', 'by', 'for', 'from', 'in', 'into', 'is', 'it', 'of', 'on', 'or', 'the',
    'to', 'with', 'please', 'can', 'could', 'would', 'should', 'this', 'that', 'there', 'here',
    'section', 'paragraph', 'clause', 'line',
    'bold', 'unbold', 'rebold', 'remove', 'make', 'set', 'change', 'format', 'formatting', 'text'
]);

function inferQueriesFromParagraphCorpus(paragraphs, userMessage, maxQueries = 8) {
    const corpus = [];
    for (const paragraph of paragraphs || []) {
        const plain = normalizeWhitespace(paragraph?.text).toLowerCase();
        const formatted = normalizeWhitespace(paragraph?.formattedText || paragraph?.text).toLowerCase();
        if (plain) corpus.push(plain);
        if (formatted && formatted !== plain) corpus.push(formatted);
    }
    if (corpus.length === 0) return [];

    const messageTokens = (String(userMessage || '').toLowerCase().match(/[a-z0-9']+/g) || [])
        .filter(token => token.length > 1)
        .filter(token => !MESSAGE_STOP_WORDS.has(token));
    if (messageTokens.length === 0) return [];

    const candidates = [];
    const seen = new Set();
    const maxN = Math.min(5, messageTokens.length);
    for (let size = maxN; size >= 2; size -= 1) {
        for (let i = 0; i <= messageTokens.length - size; i += 1) {
            const phrase = messageTokens.slice(i, i + size).join(' ');
            if (seen.has(phrase)) continue;
            seen.add(phrase);
            if (corpus.some(line => line.includes(phrase))) {
                candidates.push(phrase);
                if (candidates.length >= maxQueries) return candidates;
            }
        }
    }

    if (candidates.length > 0) return candidates;

    // Last resort: allow single-token phrase when it is distinctive and present.
    for (const token of messageTokens) {
        if (seen.has(token)) continue;
        seen.add(token);
        if (corpus.some(line => line.includes(token))) {
            candidates.push(token);
            if (candidates.length >= maxQueries) break;
        }
    }

    return candidates;
}

/**
 * Builds turn-level diagnostics for formatting-aware targeting.
 *
 * Uses quoted phrases and capitalized phrases from the user message to find
 * paragraphs where plain/formatting snapshots match.
 *
 * @param {Array<{ index?: number, text?: string, formattedText?: string }>} paragraphs
 * @param {string} userMessage
 * @param {{ maxMatches?: number }} [options]
 * @returns {{
 *   queries: string[],
 *   matches: Array<{ query: string, index: number, differs: boolean, text: string, formattedText: string }>,
 *   logLines: string[]
 * }}
 */
export function buildFormattingDiagnostics(paragraphs, userMessage, options = {}) {
    const maxMatches = Number.isInteger(options?.maxMatches) && options.maxMatches > 0
        ? options.maxMatches
        : 12;
    const items = Array.isArray(paragraphs) ? paragraphs : [];
    const queries = extractQueryPhrases(userMessage);
    if (queries.length === 0) {
        queries.push(...inferQueriesFromParagraphCorpus(items, userMessage, 8));
    }
    const matches = [];
    const logLines = [];

    if (queries.length === 0) {
        return {
            queries,
            matches,
            logLines: ['No candidate quoted/capitalized query phrases detected in user message.']
        };
    }

    for (const query of queries) {
        const queryLower = query.toLowerCase();
        for (let i = 0; i < items.length; i += 1) {
            if (matches.length >= maxMatches) break;
            const paragraph = items[i] || {};
            const index = Number.isInteger(paragraph.index) && paragraph.index > 0 ? paragraph.index : (i + 1);
            const text = normalizeWhitespace(paragraph.text);
            if (!text) continue;
            const formattedText = normalizeWhitespace(paragraph.formattedText || text) || text;
            const plainHas = text.toLowerCase().includes(queryLower);
            const formattedHas = formattedText.toLowerCase().includes(queryLower);
            if (!plainHas && !formattedHas) continue;

            const differs = formattedText !== text;
            matches.push({ query, index, differs, text, formattedText });
        }
        if (matches.length >= maxMatches) break;
    }

    if (matches.length === 0) {
        logLines.push(`No paragraph matches found for queries: ${queries.map(q => `"${q}"`).join(', ')}`);
        return { queries, matches, logLines };
    }

    for (const match of matches) {
        logLines.push(
            `Query "${match.query}" -> P${match.index} differs=${match.differs}; `
            + `plain="${truncateForLog(match.text)}"; formatted="${truncateForLog(match.formattedText)}"`
        );
    }
    if (matches.length >= maxMatches) {
        logLines.push(`Match output truncated at ${maxMatches} entries.`);
    }

    return { queries, matches, logLines };
}

/**
 * Detects whether a user message is requesting formatting removal/unformat.
 *
 * @param {string} text
 * @returns {boolean}
 */
export function isFormattingRemovalIntent(text) {
    const input = String(text || '').toLowerCase();
    if (!input.trim()) return false;
    return /(unbold|unitalic|ununderline|clear formatting|remove formatting|remove .*format|remove .*style|plain text|de-?bold|strip formatting)/i.test(input)
        || (/\bremove\b/i.test(input) && /\b(bold|italic|underline|highlight|format|formatting|style)\b/i.test(input));
}

function rankFallbackCandidate(candidate) {
    let score = 0;
    if (candidate.differs) score += 1000;
    if (candidate.source === 'assistant_quote') score += 120;
    if (candidate.source === 'user_quote') score += 100;
    if (candidate.source === 'user_query') score += 80;
    score += Math.min(40, String(candidate.phrase || '').length);
    return score;
}

/**
 * Builds a deterministic fallback redline candidate for formatting-removal
 * turns when the model returns zero operations.
 *
 * The candidate uses unchanged plain paragraph text as `modified`, which
 * triggers format-removal behavior in the reconciliation engine.
 *
 * @param {Array<{ index?: number, text?: string, formattedText?: string }>} paragraphs
 * @param {string} userMessage
 * @param {string} [assistantExplanation]
 * @returns {{ type: 'redline', targetRef: number, target: string, modified: string }|null}
 */
export function buildFormattingRemovalFallbackCandidate(paragraphs, userMessage, assistantExplanation = '') {
    const items = Array.isArray(paragraphs) ? paragraphs : [];
    if (items.length === 0) return null;

    const diagnostics = buildFormattingDiagnostics(items, userMessage);
    const phraseSources = [];
    const seen = new Set();
    const pushPhrase = (phrase, source) => {
        const normalized = normalizeWhitespace(phrase);
        if (!normalized) return;
        const key = normalized.toLowerCase();
        if (seen.has(`${source}:${key}`)) return;
        seen.add(`${source}:${key}`);
        phraseSources.push({ phrase: normalized, source });
    };

    for (const query of diagnostics.queries || []) pushPhrase(query, 'user_query');
    for (const phrase of extractQuotedPhrases(userMessage)) pushPhrase(phrase, 'user_quote');
    for (const phrase of extractQuotedPhrases(assistantExplanation)) pushPhrase(phrase, 'assistant_quote');

    const candidates = [];
    for (const sourceEntry of phraseSources) {
        const phraseLower = sourceEntry.phrase.toLowerCase();
        for (let i = 0; i < items.length; i += 1) {
            const paragraph = items[i] || {};
            const index = Number.isInteger(paragraph.index) && paragraph.index > 0 ? paragraph.index : (i + 1);
            const text = normalizeWhitespace(paragraph.text);
            if (!text) continue;
            const formattedText = normalizeWhitespace(paragraph.formattedText || text) || text;
            const plainHas = text.toLowerCase().includes(phraseLower);
            const formattedHas = formattedText.toLowerCase().includes(phraseLower);
            if (!plainHas && !formattedHas) continue;
            const differs = formattedText !== text;
            if (!differs) continue;

            candidates.push({
                phrase: sourceEntry.phrase,
                source: sourceEntry.source,
                index,
                text,
                modified: text,
                differs
            });
        }
    }

    if (candidates.length === 0 && Array.isArray(diagnostics.matches) && diagnostics.matches.length > 0) {
        for (const anchor of diagnostics.matches) {
            const anchorIndex = Number.parseInt(anchor.index, 10);
            if (!Number.isInteger(anchorIndex) || anchorIndex < 1) continue;
            const nearbyIndexes = [anchorIndex + 1, anchorIndex - 1, anchorIndex + 2, anchorIndex - 2];
            for (const idx of nearbyIndexes) {
                const paragraph = items.find(item => Number(item?.index) === idx);
                if (!paragraph) continue;
                const text = normalizeWhitespace(paragraph.text);
                if (!text) continue;
                const formattedText = normalizeWhitespace(paragraph.formattedText || text) || text;
                const differs = formattedText !== text;
                if (!differs) continue;

                candidates.push({
                    phrase: `nearby:P${anchorIndex}`,
                    source: 'section_proximity',
                    index: idx,
                    text,
                    modified: text,
                    differs
                });
            }
        }
    }

    if (candidates.length === 0) return null;
    candidates.sort((a, b) => rankFallbackCandidate(b) - rankFallbackCandidate(a));
    const best = candidates[0];
    return {
        type: 'redline',
        targetRef: best.index,
        target: best.text,
        modified: best.modified
    };
}
