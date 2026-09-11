import { applyWordOperation } from './word-operation-runner.js';
import { toScopedSharedRedlineOperation } from '@ansonlai/docx-redline-js/orchestration/redline-operation-converter.js';
import {
    preprocessMarkdown,
    ReconciliationPipeline,
    sanitizeAiResponse,
    wrapInDocumentFragment
} from '@ansonlai/docx-redline-js';
import {
    insertOoxmlWithRangeFallback,
    withNativeTrackingDisabled
} from './word-ooxml.js';

/**
 * Insert `content` (\n = paragraph break) as native, tracked Word paragraphs.
 *
 * Used for inserting into an empty paragraph or appending at end-of-document —
 * cases the OOXML reconciliation engine cannot handle (Word rejects the generated
 * package with InvalidArgument). Native insertText/insertParagraph rely on the
 * document's change-tracking mode (already set to trackAll by setChangeTrackingForAi
 * when redlines are enabled), so insertions are tracked automatically and revert
 * cleanly. Markdown is pre-rendered before native insertion so model output like
 * `# Heading` or `**Defined Term**` does not leak raw delimiters into Word.
 *
 * Bullet-list and markdown-table blocks cannot be expressed with insertText, so
 * the content is segmented (see segmentNativeInsertionBlocks) and those blocks
 * are rendered through the reconciliation engine's list/table OOXML generators
 * instead, inserted onto a native placeholder paragraph. If OOXML generation or
 * insertion fails, the block degrades to literal text with `•` markers so raw
 * markdown never leaks into the document.
 *
 * @param {Object} [opts]
 * @param {boolean} [opts.fillAnchor] - replace the anchor paragraph's (empty) text
 *   with the first line, then insert the rest after it. When false, all lines are
 *   inserted after the anchor (append).
 * @param {string} [opts.author] - redline author for generated list/table OOXML
 * @param {boolean} [opts.generateRedlines] - bake w:ins redlines into list/table OOXML
 * @param {boolean} [opts.disableNativeTracking] - toggle Word tracking off during
 *   OOXML insertion (required when redlines are baked into the OOXML)
 * @param {Word.ChangeTrackingMode|null} [opts.baseTrackingMode] - preloaded tracking mode
 */
export async function insertContentAsNativeParagraphs(context, anchorParagraph, content, opts = {}) {
    const blocks = segmentNativeInsertionBlocks(content);
    if (blocks.length === 0) return false;

    let anchor = anchorParagraph;
    let fillAnchorPending = opts.fillAnchor === true;

    for (const block of blocks) {
        if (block.kind === 'text') {
            anchor = await insertNativeTextLines(context, anchor, block.lines, fillAnchorPending);
            fillAnchorPending = false;
            continue;
        }

        // List/table block: replace a placeholder paragraph with engine-generated
        // OOXML. The sentinel paragraph is a stable native proxy to continue
        // inserting after (paragraph proxies do not survive multi-paragraph
        // insertOoxml Replace) and doubles as the spacing paragraph after the block.
        const placeholder = fillAnchorPending ? anchor : anchor.insertParagraph('', 'After');
        fillAnchorPending = false;
        const sentinel = placeholder.insertParagraph('', 'After');
        await context.sync();

        const blockText = block.lines.join('\n');
        const wrappedOoxml = block.kind === 'list'
            ? await buildListBlockFragment(blockText, opts)
            : buildTableBlockFragment(blockText, opts);

        let inserted = false;
        if (wrappedOoxml) {
            try {
                await withNativeTrackingDisabled(context, async () => {
                    await insertOoxmlWithRangeFallback(placeholder, wrappedOoxml, 'Replace', context, `NativeInsert/${block.kind}`);
                }, {
                    enabled: opts.disableNativeTracking !== false,
                    baseTrackingMode: opts.baseTrackingMode ?? null,
                    logPrefix: `NativeInsert/${block.kind}`
                });
                inserted = true;
            } catch (ooxmlError) {
                console.warn(`[NativeInsert] ${block.kind} OOXML insertion failed; falling back to literal text:`, ooxmlError);
            }
        }
        if (!inserted) {
            const literalLines = block.kind === 'list'
                ? block.lines.map(bulletLineToLiteralText)
                : block.lines;
            await insertNativeTextLines(context, placeholder, literalLines, true);
        }
        anchor = sentinel;
    }
    return true;
}

const BULLET_LINE_RE = /^\s*[-*+•]\s+/;
// Nested (indented) numbered/alpha/roman items are allowed to continue a bullet
// block; nested numbering restarts at 1 by convention so engine generation is safe.
const INDENTED_MARKER_LINE_RE = /^\s+(?:\d+(?:\.\d+)*[.)]|[A-Za-z][.)]|[ivxlcIVXLC]+[.)])\s+/;

/**
 * Splits fallback insertion content into text / list / table blocks.
 *
 * Only bullet-marker lines start a list block: the engine's generated numbered
 * lists always restart at 1, so top-level numbered lines ("3. Exclusions") must
 * stay literal text to preserve the model's explicit numbering.
 *
 * A single blank source line immediately after a list/table block is consumed:
 * the sentinel/spacing paragraph inserted after the block already represents it.
 *
 * @param {string} content
 * @returns {Array<{ kind: 'text'|'list'|'table', lines: string[] }>}
 */
export function segmentNativeInsertionBlocks(content) {
    const lines = String(content == null ? '' : content).split('\n');
    const blocks = [];
    const pushTextLine = (line) => {
        const last = blocks[blocks.length - 1];
        if (last && last.kind === 'text') {
            last.lines.push(line);
        } else {
            blocks.push({ kind: 'text', lines: [line] });
        }
    };
    const consumeTrailingBlank = (index) => (
        index < lines.length && lines[index].trim() === '' ? index + 1 : index
    );

    let i = 0;
    while (i < lines.length) {
        const line = lines[i];

        if (isMarkdownTableStart(lines, i)) {
            const tableLines = [];
            while (i < lines.length && /^\s*\|/.test(lines[i])) {
                tableLines.push(lines[i]);
                i += 1;
            }
            blocks.push({ kind: 'table', lines: tableLines });
            i = consumeTrailingBlank(i);
            continue;
        }

        if (BULLET_LINE_RE.test(line)) {
            const listLines = [];
            while (i < lines.length && (BULLET_LINE_RE.test(lines[i]) || INDENTED_MARKER_LINE_RE.test(lines[i]))) {
                listLines.push(lines[i]);
                i += 1;
            }
            blocks.push({ kind: 'list', lines: listLines });
            i = consumeTrailingBlank(i);
            continue;
        }

        pushTextLine(line);
        i += 1;
    }

    return blocks;
}

function isMarkdownTableStart(lines, index) {
    return /^\s*\|.*\|\s*$/.test(lines[index] || '')
        && isMarkdownSeparatorRow(lines[index + 1] || '');
}

function bulletLineToLiteralText(line) {
    return String(line || '').replace(/^(\s*)[-*+]\s+/, '$1• ');
}

async function insertNativeTextLines(context, startParagraph, lines, replaceFirstIntoAnchor) {
    let anchor = startParagraph;
    let replacePending = replaceFirstIntoAnchor === true;
    for (const line of lines) {
        const prepared = prepareNativeMarkdownParagraph(line);
        if (replacePending) {
            anchor.insertText(prepared.text, 'Replace');
            replacePending = false;
        } else {
            anchor = anchor.insertParagraph(prepared.text, 'After');
        }
        await context.sync();
        await applyNativeParagraphFormatting(anchor, prepared, context);
    }
    return anchor;
}

async function buildListBlockFragment(blockText, opts) {
    try {
        const pipeline = new ReconciliationPipeline({
            generateRedlines: opts.generateRedlines !== false,
            author: opts.author,
            font: opts.font || null
        });
        const result = await pipeline.executeListGeneration(blockText, null, null, '');
        if (!result?.ooxml || result.isValid === false) {
            console.warn('[NativeInsert] List generation produced no valid OOXML:', result?.warnings);
            return null;
        }
        return wrapInDocumentFragment(result.ooxml, {
            includeNumbering: true,
            numberingXml: result.numberingXml
        });
    } catch (listError) {
        console.warn('[NativeInsert] List OOXML generation failed:', listError);
        return null;
    }
}

function buildTableBlockFragment(blockText, opts) {
    try {
        const pipeline = new ReconciliationPipeline({
            generateRedlines: opts.generateRedlines !== false,
            author: opts.author
        });
        const result = pipeline.executeTableGeneration(blockText);
        if (!result?.ooxml || result.isValid === false) {
            console.warn('[NativeInsert] Table generation produced no valid OOXML:', result?.warnings);
            return null;
        }
        return wrapInDocumentFragment(result.ooxml, { includeNumbering: false });
    } catch (tableError) {
        console.warn('[NativeInsert] Table OOXML generation failed:', tableError);
        return null;
    }
}

export function prepareNativeMarkdownParagraphs(content) {
    return String(content == null ? '' : content)
        .split('\n')
        .map(prepareNativeMarkdownParagraph);
}

function prepareNativeMarkdownParagraph(line) {
    const raw = String(line == null ? '' : line);
    const headingMatch = raw.match(/^\s{0,3}(#{1,9})\s+(.+?)\s*#*\s*$/);
    const headingLevel = headingMatch ? Math.min(headingMatch[1].length, 9) : null;
    const markdownText = headingMatch ? headingMatch[2] : raw;
    const { cleanText, formatHints } = preprocessMarkdown(markdownText);
    const hints = Array.isArray(formatHints) ? [...formatHints] : [];

    if (headingLevel && cleanText.trim()) {
        hints.push({
            start: 0,
            end: cleanText.length,
            format: { bold: true }
        });
    }

    return {
        text: cleanText,
        formatHints: hints,
        headingLevel
    };
}

async function applyNativeParagraphFormatting(paragraph, prepared, context) {
    if (!paragraph || !prepared) return;

    if (prepared.headingLevel) {
        try {
            paragraph.style = `Heading ${prepared.headingLevel}`;
        } catch {
            // Some Word hosts reject style assignment on freshly inserted proxies.
        }
        if (paragraph.font) {
            paragraph.font.bold = true;
        }
        await context.sync();
    }

    await applyNativeFormatHints(paragraph, prepared.text, prepared.formatHints, context);
}

async function applyNativeFormatHints(paragraph, text, formatHints, context) {
    if (!paragraph || typeof paragraph.getRange !== 'function') return;
    if (!Array.isArray(formatHints) || formatHints.length === 0) return;

    const paragraphRange = paragraph.getRange();
    const usedByText = new Map();

    for (const hint of formatHints) {
        const hintText = String(text || '').substring(hint.start, hint.end);
        if (!hintText || hintText.trim().length === 0) continue;

        try {
            const searchResults = paragraphRange.search(hintText, {
                matchCase: true,
                matchWholeWord: false
            });
            searchResults.load('items/text');
            await context.sync();

            const items = Array.isArray(searchResults.items) ? searchResults.items : [];
            if (items.length === 0) continue;

            const used = usedByText.get(hintText) || 0;
            const targetRange = items[Math.min(used, items.length - 1)];
            usedByText.set(hintText, used + 1);

            if (hint.format?.bold) {
                targetRange.font.bold = true;
            }
            if (hint.format?.italic) {
                targetRange.font.italic = true;
            }
            if (hint.format?.underline) {
                targetRange.font.underline = globalThis.Word?.UnderlineType?.single || 'Single';
            }
            if (hint.format?.strikethrough) {
                targetRange.font.strikeThrough = true;
            }
            await context.sync();
        } catch (formatError) {
            console.warn(`Could not apply native markdown formatting to "${hintText}":`, formatError);
        }
    }
}
import { isMarkdownTableText } from '@ansonlai/docx-redline-js/core/paragraph-targeting.js';
import { isLikelyStructuredTableSourceParagraph } from '@ansonlai/docx-redline-js/core/table-targeting.js';

function normalizeNeedleText(value) {
    if (value == null) return '';
    return String(value)
        .replace(/\\n/g, '\n')
        .replace(/\\t/g, '\t')
        .replace(/\\r/g, '\r')
        .trim();
}

function normalizeReplacementText(value) {
    if (value == null) return '';
    return String(value)
        .replace(/\\n/g, '\n')
        .replace(/\\t/g, '\t')
        .replace(/\\r/g, '\r');
}

function getReplacementText(change, operationName) {
    if (operationName === 'edit_paragraph') {
        return change?.newContent;
    }
    if (operationName === 'replace_paragraph' || operationName === 'replace_range') {
        return change?.content ?? change?.newContent ?? change?.replacementText;
    }
    if (operationName === 'modify_text') {
        return change?.replacementText;
    }
    return null;
}

function hasStructuralReplacementField(change) {
    if (!change || typeof change !== 'object') return false;
    return Object.prototype.hasOwnProperty.call(change, 'content')
        || Object.prototype.hasOwnProperty.call(change, 'newContent');
}

function setReplacementText(change, operationName, replacementText) {
    if (operationName === 'edit_paragraph') {
        change.newContent = replacementText;
        return;
    }
    if (operationName === 'replace_paragraph' || operationName === 'replace_range') {
        change.content = replacementText;
    }
}

function splitPipeCells(line) {
    return String(line || '')
        .split('|')
        .map(cell => cell.trim())
        .filter(cell => cell.length > 0);
}

function stripTrailingColon(text) {
    return String(text || '').trim().replace(/:\s*$/, '');
}

function markdownRow(cells) {
    return `| ${cells.map(cell => String(cell || '').trim()).join(' | ')} |`;
}

function markdownSeparator(columnCount) {
    return markdownRow(new Array(Math.max(1, columnCount)).fill('---'));
}

function markdownTableForLabeledColumns(columns) {
    const safeColumns = Array.isArray(columns)
        ? columns.filter(column => column && column.label)
        : [];
    if (safeColumns.length < 2) return null;

    const maxDetailRows = safeColumns.reduce(
        (max, column) => Math.max(max, Array.isArray(column.details) ? column.details.length : 0),
        0
    );

    const bodyRows = [];
    for (let rowIndex = 0; rowIndex < Math.max(1, maxDetailRows); rowIndex += 1) {
        bodyRows.push(markdownRow(
            safeColumns.map(column => String(column.details?.[rowIndex] || '').trim())
        ));
    }

    return [
        markdownRow(safeColumns.map(column => column.label)),
        markdownSeparator(safeColumns.length),
        ...bodyRows
    ].join('\n');
}

function isMarkdownSeparatorRow(line) {
    const normalized = String(line || '').replace(/\s+/g, '');
    return /^\|:?-{3,}:?(\|:?-{3,}:?)+\|?$/.test(normalized);
}

function splitCellBreaks(cell) {
    return String(cell || '')
        .split(/<br\s*\/?>/gi)
        .map(part => part.trim());
}

function expandMarkdownTableHtmlBreaks(text) {
    const lines = normalizeReplacementText(text)
        .trim()
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(line => line.startsWith('|'));
    if (lines.length === 0 || !lines.some(line => /<br\s*\/?>/i.test(line))) {
        return normalizeReplacementText(text).trim();
    }

    const expanded = [];
    for (const line of lines) {
        if (isMarkdownSeparatorRow(line)) {
            expanded.push(line);
            continue;
        }

        const cells = splitPipeCells(line);
        const splitCells = cells.map(splitCellBreaks);
        const rowCount = splitCells.reduce((max, parts) => Math.max(max, parts.length), 1);
        for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
            expanded.push(markdownRow(splitCells.map(parts => parts[rowIndex] || '')));
        }
    }

    return expanded.join('\n');
}

function normalizeInlineLabeledBlockTable(cells) {
    const groups = [];
    let current = [];

    for (const cell of cells) {
        if (isTableRangeConnectorText(cell)) {
            if (current.length > 0) {
                groups.push(current);
                current = [];
            }
            continue;
        }
        current.push(cell);
    }
    if (current.length > 0) {
        groups.push(current);
    }

    if (groups.length < 2 || groups.some(group => group.length < 2)) {
        return null;
    }

    const columns = groups.map(group => ({
        label: stripTrailingColon(group[0]),
        details: group.slice(1)
    }));
    if (columns.some(column => !column.label || /[.!?]$/.test(column.label))) {
        return null;
    }

    return markdownTableForLabeledColumns(columns);
}

function normalizeInlinePipeTable(text) {
    const cells = splitPipeCells(text);
    if (cells.length < 2) {
        return null;
    }

    const labeledBlockTable = normalizeInlineLabeledBlockTable(cells);
    if (labeledBlockTable) {
        return labeledBlockTable;
    }

    const columnCount = cells.length % 2 === 0 ? 2 : Math.min(3, cells.length);
    if (columnCount < 2 || cells.length < columnCount) {
        return null;
    }

    const rows = [];
    for (let i = 0; i < cells.length; i += columnCount) {
        const row = cells.slice(i, i + columnCount);
        while (row.length < columnCount) row.push('');
        rows.push(row);
    }

    return [
        markdownRow(rows[0]),
        markdownSeparator(columnCount),
        ...rows.slice(1).map(markdownRow)
    ].join('\n');
}

export function normalizeTableReplacementText(text) {
    const normalized = normalizeReplacementText(text).trim();
    if (!normalized) {
        return normalized;
    }
    if (isMarkdownTableText(normalized)) {
        return expandMarkdownTableHtmlBreaks(normalized);
    }

    const lines = normalized
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(Boolean);

    if (lines.length === 1 && lines[0].includes('|')) {
        return normalizeInlinePipeTable(lines[0]) || normalized;
    }

    const pipeLines = lines.filter(line => line.includes('|'));
    if (pipeLines.length >= 2) {
        const rows = pipeLines.map(splitPipeCells).filter(row => row.length > 0);
        const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
        if (columnCount >= 2) {
            const normalizedRows = rows.map(row => {
                const copy = row.slice(0, columnCount);
                while (copy.length < columnCount) copy.push('');
                return copy;
            });
            return [
                markdownRow(normalizedRows[0]),
                markdownSeparator(columnCount),
                ...normalizedRows.slice(1).map(markdownRow)
            ].join('\n');
        }
    }

    return normalized;
}

function replacementLooksTableLike(text) {
    const normalized = normalizeReplacementText(text).trim();
    return isMarkdownTableText(normalized)
        || (normalized.includes('|') && splitPipeCells(normalized).length >= 2);
}

export function isDiagnosticReplacementText(text) {
    const normalized = normalizeReplacementText(text)
        .replace(/\s+/g, ' ')
        .trim()
        .toLowerCase();
    if (!normalized) return false;

    const diagnosticPatterns = [
        /schema requires/,
        /schema.*paragraphindex/,
        /replace_range operation/,
        /replace_paragraph operation/,
        /do not include ["']?originaltext/,
        /do not include .*replacementtext/,
        /originaltext.*replacementtext/,
        /provide only .*paragraphindex.*operation.*content/,
        /none required for replace_range/,
        /invalid replace_range/
    ];

    return diagnosticPatterns.some(pattern => pattern.test(normalized));
}

function isLikelyTableSourceText(text) {
    const normalized = String(text || '').trim();
    if (!normalized) return false;
    if (isTableRangeConnectorText(normalized)) return true;
    if (isLikelyStructuredTableSourceParagraph(normalized)) return true;

    const logicalLines = normalized
        .split(/[\r\n\v]+/)
        .map(line => line.trim())
        .filter(Boolean);

    return logicalLines.length > 1
        && logicalLines.some(line => /:\s*$/.test(line))
        && logicalLines.some(line => /^\[.*\]$/.test(line) || /^\(.*\)$/.test(line));
}

function getLogicalParagraphLines(text) {
    return String(text || '')
        .split(/[\r\n\v]+/)
        .map(line => line.trim())
        .filter(Boolean);
}

function isTableRangeConnectorText(text) {
    return /^(?:and|or|between|with|plus|to)$/i.test(String(text || '').trim());
}

function parseLabeledBlockParagraph(text) {
    const lines = getLogicalParagraphLines(text);
    if (lines.length < 2) return null;

    const firstLine = lines[0];
    const label = stripTrailingColon(firstLine);
    const hasExplicitLabel = /:\s*$/.test(firstLine);
    const looksLikeCompactHeading = label.length > 0
        && label.length <= 80
        && !/[.!?]$/.test(label)
        && !/\s(?:is|are|was|were|shall|must|may|will|can|could|should|would)\s/i.test(label);

    if (!hasExplicitLabel && !looksLikeCompactHeading) return null;

    return {
        label,
        details: lines.slice(1)
    };
}

export function synthesizeMarkdownTableFromSourceRange(paragraphItems, startIndex, endIndex) {
    const items = Array.isArray(paragraphItems) ? paragraphItems : [];
    if (
        !Number.isInteger(startIndex)
        || !Number.isInteger(endIndex)
        || startIndex < 0
        || endIndex < startIndex
        || endIndex >= items.length
    ) {
        return null;
    }

    const columns = [];
    for (let index = startIndex; index <= endIndex; index += 1) {
        const text = String(items[index]?.text || '').trim();
        if (!text || isTableRangeConnectorText(text)) continue;

        const parsed = parseLabeledBlockParagraph(text);
        if (parsed) {
            columns.push(parsed);
        }
    }

    if (columns.length >= 2) {
        return markdownTableForLabeledColumns(columns);
    }

    return null;
}

export function inferTableConversionEndIndex(paragraphItems, startIndex, maxScan = 10) {
    const items = Array.isArray(paragraphItems) ? paragraphItems : [];
    if (!Number.isInteger(startIndex) || startIndex < 0 || startIndex >= items.length) {
        return startIndex;
    }

    let endIndex = startIndex;
    const limit = Math.min(items.length - 1, startIndex + Math.max(1, maxScan));
    for (let index = startIndex + 1; index <= limit; index += 1) {
        const text = String(items[index]?.text || '').trim();
        if (!text) {
            if (endIndex > startIndex) break;
            continue;
        }
        if (!isLikelyTableSourceText(text)) {
            break;
        }
        endIndex = index;
    }

    return endIndex;
}

export function findNearbyParagraphIndexForModifyText(paragraphItems, startIndex, change) {
    const items = Array.isArray(paragraphItems) ? paragraphItems : [];
    if (!Number.isInteger(startIndex) || startIndex < 0 || startIndex >= items.length) {
        return startIndex;
    }

    const originalText = normalizeNeedleText(change?.originalText);
    if (!originalText) {
        return startIndex;
    }

    const textAt = index => String(items[index]?.text ?? '');
    const startText = textAt(startIndex);
    if (startText.trim().length > 0) {
        return startIndex;
    }

    const originalLower = originalText.toLowerCase();
    const maxDistance = 12;
    let bestExact = null;
    let bestInsensitive = null;

    for (let distance = 1; distance <= maxDistance; distance += 1) {
        const candidates = [startIndex - distance, startIndex + distance];
        for (const candidateIndex of candidates) {
            if (candidateIndex < 0 || candidateIndex >= items.length) continue;

            const candidateText = textAt(candidateIndex);
            if (!candidateText.trim()) continue;

            if (candidateText.includes(originalText)) {
                if (bestExact == null) {
                    bestExact = candidateIndex;
                }
                continue;
            }

            if (candidateText.toLowerCase().includes(originalLower)) {
                if (bestInsensitive == null) {
                    bestInsensitive = candidateIndex;
                }
            }
        }

        if (bestExact != null) return bestExact;
    }

    if (bestInsensitive != null) return bestInsensitive;
    return startIndex;
}

function parseParagraphIndex(value) {
    const parsed = Number.parseInt(String(value ?? ''), 10);
    return Number.isInteger(parsed) ? parsed - 1 : null;
}

export async function applyRedlineChangesToWordContext(context, aiChanges, options = {}) {
    const sourceChanges = Array.isArray(aiChanges) ? aiChanges : [];
    let sanitizedFieldCount = 0;
    const changes = options.sanitizeInput === true
        ? sourceChanges.map(change => {
            if (!change || typeof change !== 'object') return change;
            const sanitizedChange = { ...change };
            for (const field of ['content', 'newContent', 'replacementText']) {
                if (typeof sanitizedChange[field] !== 'string') continue;
                const sanitized = sanitizeAiResponse(sanitizedChange[field]);
                if (sanitized !== sanitizedChange[field]) sanitizedFieldCount += 1;
                sanitizedChange[field] = sanitized;
            }
            return sanitizedChange;
        })
        : sourceChanges;
    const logPrefix = options.logPrefix || 'Redline/Shared';
    const onInfo = typeof options.onInfo === 'function'
        ? options.onInfo
        : message => console.log(`[${logPrefix}] ${message}`);
    const onWarn = typeof options.onWarn === 'function'
        ? options.onWarn
        : message => console.warn(`[${logPrefix}] ${message}`);
    if (sanitizedFieldCount > 0) {
        onWarn(`Sanitized ${sanitizedFieldCount} model-generated document content field(s).`);
    }
    const requestedContentKind = options.requestedContentKind || null;
    const requiresTableContent = requestedContentKind === 'table';

    let changesApplied = 0;
    const skipped = [];
    const recordSkip = (change, operation, reason, code = undefined, details = undefined) => {
        skipped.push({
            paragraphIndex: change?.paragraphIndex,
            operation: operation || String(change?.operation || '').trim().toLowerCase(),
            reason,
            ...(code ? { code } : {}),
            ...(details?.receipt ? { receipt: details.receipt } : {}),
            ...(details?.rolledBack !== undefined ? { rolledBack: details.rolledBack } : {})
        });
    };

    for (const change of changes) {
        try {
            let operationName = String(change?.operation || '').trim().toLowerCase();
            const startIndex = parseParagraphIndex(change?.paragraphIndex);
            if (!Number.isInteger(startIndex) || startIndex < 0) {
                onWarn(`Invalid start paragraph index: ${change?.paragraphIndex}`);
                continue;
            }

            if (requiresTableContent && operationName === 'modify_text') {
                onWarn(
                    `Skipping modify_text at P${change?.paragraphIndex}: table conversion requests require replace_paragraph or replace_range with markdown table content.`
                );
                continue;
            }

            const paragraphs = context.document.body.paragraphs;
            paragraphs.load('items/text');
            await context.sync();

            const paragraphCount = paragraphs.items.length;

            // Append-at-end: targeting the paragraph one past the last one with a
            // content-bearing operation inserts brand-new paragraph(s) after the
            // document's last paragraph (the only way to add content when there is
            // no trailing blank paragraph to reuse).
            const isAppendAtEnd = startIndex === paragraphCount
                && (operationName === 'replace_paragraph' || operationName === 'replace_range' || operationName === 'edit_paragraph');

            if (startIndex > paragraphCount || (startIndex === paragraphCount && !isAppendAtEnd)) {
                onWarn(`Out-of-range target P${change?.paragraphIndex} (count=${paragraphCount}); no-op.`);
                recordSkip(change, operationName, `paragraph P${change?.paragraphIndex} is out of range (document has ${paragraphCount} paragraphs)`);
                continue;
            }

            if (isAppendAtEnd) {
                const appendContent = normalizeReplacementText(getReplacementText(change, operationName));
                if (!appendContent || !appendContent.trim()) {
                    onWarn(`Append at P${change?.paragraphIndex}: no content to append; no-op.`);
                    recordSkip(change, operationName, 'no content was provided to append at the end of the document');
                    continue;
                }
                if (paragraphCount === 0) {
                    onWarn('Append requested but document has no paragraphs; no-op.');
                    recordSkip(change, operationName, 'document has no paragraphs to append after');
                    continue;
                }
                // Append new paragraphs after the last one using native (tracked) APIs.
                const lastParagraph = paragraphs.items[paragraphCount - 1];
                await insertContentAsNativeParagraphs(context, lastParagraph, appendContent, {
                    author: options.author,
                    generateRedlines: options.generateRedlines,
                    disableNativeTracking: options.disableNativeTracking,
                    baseTrackingMode: options.baseTrackingMode ?? null
                });
                changesApplied += 1;
                onInfo(`Appended new content after P${paragraphCount} (end of document).`);
                continue;
            }

            let effectiveStartIndex = startIndex;
            if (operationName === 'modify_text') {
                effectiveStartIndex = findNearbyParagraphIndexForModifyText(paragraphs.items, startIndex, change);
                if (effectiveStartIndex !== startIndex) {
                    onWarn(
                        `Rebased modify_text from P${startIndex + 1} to P${effectiveStartIndex + 1} based on originalText match.`
                    );
                }
            }

            const normalizedChange = { ...change, operation: operationName };
            if (operationName === 'edit_paragraph' || operationName === 'replace_paragraph' || operationName === 'replace_range') {
                const replacementText = getReplacementText(normalizedChange, operationName);
                const normalizedTableText = normalizeTableReplacementText(replacementText);
                if (replacementLooksTableLike(replacementText) && normalizedTableText !== normalizeReplacementText(replacementText).trim()) {
                    setReplacementText(normalizedChange, operationName, normalizedTableText);
                    onWarn('Normalized pipe-delimited table replacement into markdown table syntax.');
                }

                if (isMarkdownTableText(normalizedTableText) && operationName === 'edit_paragraph') {
                    operationName = 'replace_paragraph';
                    normalizedChange.operation = operationName;
                    normalizedChange.content = normalizedTableText;
                    delete normalizedChange.newContent;
                    onWarn('Promoted table edit_paragraph to replace_paragraph for OOXML table generation.');
                }
            }

            let endIndex = effectiveStartIndex;
            let insertionBeforeStart = false;
            if (operationName === 'replace_range') {
                let requestedEndIndex = parseParagraphIndex(normalizedChange?.endParagraphIndex);
                if (!Number.isInteger(requestedEndIndex) || requestedEndIndex < -1) {
                    if (requiresTableContent || replacementLooksTableLike(getReplacementText(normalizedChange, operationName))) {
                        requestedEndIndex = inferTableConversionEndIndex(paragraphs.items, effectiveStartIndex);
                        normalizedChange.endParagraphIndex = requestedEndIndex + 1;
                        onWarn(
                            `Inferred missing table replace_range end as P${requestedEndIndex + 1}.`
                        );
                    } else {
                        onWarn(`Invalid replace_range endParagraphIndex: ${normalizedChange?.endParagraphIndex}; no-op.`);
                        continue;
                    }
                }

                if (requestedEndIndex === effectiveStartIndex - 1) {
                    insertionBeforeStart = true;
                    endIndex = effectiveStartIndex;
                    onWarn(
                        `Normalizing replace_range insertion-before-target (P${effectiveStartIndex + 1}..P${requestedEndIndex + 1}) to scoped insertion.`
                    );
                } else {
                    endIndex = requestedEndIndex;
                    if (endIndex < effectiveStartIndex || endIndex >= paragraphCount) {
                        onWarn(`Invalid replace_range endParagraphIndex: ${normalizedChange?.endParagraphIndex}; no-op.`);
                        continue;
                    }
                }

                if (!hasStructuralReplacementField(normalizedChange)) {
                    const replacementCandidate = getReplacementText(normalizedChange, operationName);
                    const shouldAttemptTableRecovery = requiresTableContent
                        || isDiagnosticReplacementText(replacementCandidate);
                    const synthesizedTable = synthesizeMarkdownTableFromSourceRange(
                        paragraphs.items,
                        effectiveStartIndex,
                        endIndex
                    );
                    if (shouldAttemptTableRecovery && synthesizedTable) {
                        normalizedChange.content = synthesizedTable;
                        onWarn('Synthesized missing table content from the source paragraph range.');
                    } else {
                        onWarn(
                            `Skipping malformed replace_range at P${effectiveStartIndex + 1}: missing required content.`
                        );
                        continue;
                    }
                }
            } else if (
                operationName === 'replace_paragraph'
                && (requiresTableContent || replacementLooksTableLike(getReplacementText(normalizedChange, operationName)))
            ) {
                const inferredEndIndex = inferTableConversionEndIndex(paragraphs.items, effectiveStartIndex);
                if (inferredEndIndex > effectiveStartIndex) {
                    operationName = 'replace_range';
                    normalizedChange.operation = operationName;
                    normalizedChange.endParagraphIndex = inferredEndIndex + 1;
                    endIndex = inferredEndIndex;
                    onWarn(
                        `Expanded table replacement scope from P${effectiveStartIndex + 1} to P${inferredEndIndex + 1}.`
                    );
                }
            }

            const replacementAfterRangeResolution = getReplacementText(normalizedChange, operationName);
            if (
                (operationName === 'replace_paragraph' || operationName === 'replace_range')
                && isDiagnosticReplacementText(replacementAfterRangeResolution)
            ) {
                const synthesizedTable = synthesizeMarkdownTableFromSourceRange(
                    paragraphs.items,
                    effectiveStartIndex,
                    endIndex
                );
                if (synthesizedTable) {
                    normalizedChange.content = synthesizedTable;
                    onWarn('Replaced diagnostic model output with synthesized table content from the source range.');
                } else {
                    onWarn(
                        `Skipping malformed ${operationName} at P${effectiveStartIndex + 1}: replacement text contains model diagnostics.`
                    );
                    continue;
                }
            }

            if (
                (operationName === 'replace_paragraph' || operationName === 'replace_range')
                && !hasStructuralReplacementField(normalizedChange)
            ) {
                onWarn(
                    `Skipping malformed ${operationName} at P${effectiveStartIndex + 1}: missing required content.`
                );
                continue;
            }

            if (
                requiresTableContent
                && (operationName === 'edit_paragraph' || operationName === 'replace_paragraph' || operationName === 'replace_range')
            ) {
                const finalReplacementText = getReplacementText(normalizedChange, operationName);
                const normalizedTableText = normalizeTableReplacementText(finalReplacementText);
                if (
                    replacementLooksTableLike(finalReplacementText)
                    && normalizedTableText !== normalizeReplacementText(finalReplacementText).trim()
                ) {
                    setReplacementText(normalizedChange, operationName, normalizedTableText);
                    onWarn('Normalized final table replacement into markdown table syntax.');
                }

                const finalContent = getReplacementText(normalizedChange, operationName);
                if (!isMarkdownTableText(normalizeReplacementText(finalContent).trim())) {
                    const synthesizedTable = synthesizeMarkdownTableFromSourceRange(
                        paragraphs.items,
                        effectiveStartIndex,
                        endIndex
                    );
                    if (synthesizedTable) {
                        operationName = 'replace_range';
                        normalizedChange.operation = operationName;
                        normalizedChange.content = synthesizedTable;
                        normalizedChange.endParagraphIndex = endIndex + 1;
                        delete normalizedChange.newContent;
                        onWarn('Replaced non-table model output with synthesized table content from the source range.');
                    } else {
                        onWarn(
                            `Skipping ${operationName} at P${effectiveStartIndex + 1}: table request did not produce markdown table content.`
                        );
                        continue;
                    }
                }
            }

            const startParagraph = paragraphs.items[effectiveStartIndex];
            const scopeParagraphCount = insertionBeforeStart
                ? 1
                : (endIndex - effectiveStartIndex) + 1;
            const sourceParagraphText = String(startParagraph.text || '');
            const emptySingleTarget = !insertionBeforeStart
                && scopeParagraphCount === 1
                && sourceParagraphText.trim().length === 0;

            if (
                emptySingleTarget
                && (
                    operationName === 'edit_paragraph'
                    || operationName === 'replace_paragraph'
                    || operationName === 'replace_range'
                )
            ) {
                const insertContent = normalizeReplacementText(getReplacementText(normalizedChange, operationName));
                if (insertContent.trim()) {
                    onInfo(`Empty target P${change?.paragraphIndex}: inserting content as a native tracked change before redline conversion.`);
                    await insertContentAsNativeParagraphs(context, startParagraph, insertContent, {
                        fillAnchor: true,
                        author: options.author,
                        generateRedlines: options.generateRedlines,
                        disableNativeTracking: options.disableNativeTracking,
                        baseTrackingMode: options.baseTrackingMode ?? null
                    });
                    changesApplied += 1;
                    continue;
                }
            }

            const converted = toScopedSharedRedlineOperation(normalizedChange, {
                scopeStartText: sourceParagraphText,
                scopeParagraphCount,
                insertionBeforeStart
            });
            if (!converted.ok) {
                // Empty target paragraph: the OOXML reconciliation engine cannot diff
                // against empty text (Word rejects its output). Insert the content with
                // native, tracked Word APIs instead of skipping.
                const emptyTarget = /target.*empty|empty.*target/i.test(converted.reason || '');
                const insertContent = emptyTarget
                    ? normalizeReplacementText(getReplacementText(normalizedChange, operationName))
                    : null;
                if (insertContent && insertContent.trim()) {
                    onInfo(`Empty target P${change?.paragraphIndex}: inserting content as a native tracked change.`);
                    await insertContentAsNativeParagraphs(context, startParagraph, insertContent, {
                        fillAnchor: true,
                        author: options.author,
                        generateRedlines: options.generateRedlines,
                        disableNativeTracking: options.disableNativeTracking,
                        baseTrackingMode: options.baseTrackingMode ?? null
                    });
                    changesApplied += 1;
                    continue;
                }
                onWarn(`Skipping change: ${converted.reason}`);
                recordSkip(change, operationName, emptyTarget ? 'empty paragraph and no content to insert' : converted.reason);
                continue;
            }

            const applied = await applyWordOperation(
                context,
                converted.operation,
                scopeParagraphCount === 1
                    ? { paragraph: startParagraph }
                    : { paragraph: startParagraph, endParagraph: paragraphs.items[endIndex] },
                {
                    author: options.author,
                    generateRedlines: options.generateRedlines,
                    sanitizeInput: options.sanitizeInput === true,
                    disableNativeTracking: options.disableNativeTracking,
                    baseTrackingMode: options.baseTrackingMode ?? null,
                    logPrefix,
                    onInfo,
                    onWarn
                }
            );

            if (applied) {
                changesApplied += 1;
            } else {
                onWarn(`No changes produced for change: ${JSON.stringify(normalizedChange)}`);
                recordSkip(change, operationName, `no changes were produced for P${change?.paragraphIndex} (the new content may match the existing text)`);
            }
        } catch (changeError) {
            onWarn(`Failed to apply change ${JSON.stringify(change)}: ${changeError?.message || changeError}`);
            recordSkip(
                change,
                undefined,
                `apply error: ${changeError?.message || changeError}`,
                changeError?.code,
                changeError?.details
            );
        }
    }

    onInfo(`Total changes applied: ${changesApplied}`);
    return { changesApplied, skipped };
}
