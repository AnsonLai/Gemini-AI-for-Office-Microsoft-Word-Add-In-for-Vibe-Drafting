// Host-independent scoring for the golden scenario. Reads an exported .docx,
// resolves the accepted view with the shared library, and evaluates declarative
// checks against paragraphs, effective run formatting, lists, tables and comments.
import { openDocx } from '@ansonlai/docx-redline-js';
import { unzipDocx } from '@ansonlai/docx-redline-js/document/zip-archive.js';
import { parseOoxmlSafe } from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const decoder = new TextDecoder();
const FORMAT_KEYS = ['bold', 'italic', 'underline', 'strikethrough', 'highlight'];

const children = (node, name) => Array.from(node?.childNodes || [])
    .filter(child => child.nodeType === 1 && child.localName === name);
const child = (node, name) => children(node, name)[0] || null;
const descendants = (node, name) => node ? Array.from(node.getElementsByTagNameNS(NS_W, name)) : [];
const attr = (node, name) => node?.getAttributeNS(NS_W, name) ?? node?.getAttribute(`w:${name}`) ?? null;

function parse(xml, part) {
    const parsed = parseOoxmlSafe(xml);
    if (!parsed.doc) throw new Error(`${part} is not parseable XML: ${parsed.error?.message || 'unknown error'}`);
    return parsed.doc;
}

function toggleValue(element) {
    if (!element) return undefined;
    const value = attr(element, 'val');
    return value == null || !['0', 'false', 'off', 'none'].includes(String(value).toLowerCase());
}

/** Formatting declared directly in one w:rPr (undefined = not declared). */
function declaredFormat(rPr) {
    if (!rPr) return {};
    const format = {};
    const bold = toggleValue(child(rPr, 'b'));
    if (bold !== undefined) format.bold = bold;
    const italic = toggleValue(child(rPr, 'i'));
    if (italic !== undefined) format.italic = italic;
    const strike = toggleValue(child(rPr, 'strike'));
    if (strike !== undefined) format.strikethrough = strike;
    const underline = child(rPr, 'u');
    if (underline) format.underline = String(attr(underline, 'val') || 'single').toLowerCase() !== 'none';
    const highlight = child(rPr, 'highlight');
    if (highlight) format.highlight = String(attr(highlight, 'val') || 'none').toLowerCase() !== 'none';
    const shading = child(rPr, 'shd');
    if (shading && format.highlight === undefined) {
        const fill = String(attr(shading, 'fill') || 'auto').toLowerCase();
        if (fill !== 'auto' && fill !== 'ffffff') format.highlight = true;
    }
    return format;
}

function buildStyles(stylesXml) {
    const styles = new Map();
    if (!stylesXml) return styles;
    const doc = parse(stylesXml, 'word/styles.xml');
    for (const style of descendants(doc, 'style')) {
        styles.set(attr(style, 'styleId'), {
            basedOn: attr(child(style, 'basedOn'), 'val'),
            format: declaredFormat(child(style, 'rPr')),
            numPr: child(child(style, 'pPr'), 'numPr')
        });
    }
    return styles;
}

function styleFormat(styles, styleId, seen = new Set()) {
    if (!styleId || seen.has(styleId) || !styles.has(styleId)) return {};
    seen.add(styleId);
    const style = styles.get(styleId);
    return { ...styleFormat(styles, style.basedOn, seen), ...style.format };
}

function buildNumbering(numberingXml) {
    const formats = new Map();
    if (!numberingXml) return formats;
    const doc = parse(numberingXml, 'word/numbering.xml');
    const abstracts = new Map(descendants(doc, 'abstractNum').map(abstract => [
        attr(abstract, 'abstractNumId'),
        new Map(children(abstract, 'lvl').map(level => [Number(attr(level, 'ilvl')), {
            numFmt: attr(child(level, 'numFmt'), 'val'),
            lvlText: attr(child(level, 'lvlText'), 'val')
        }]))
    ]));
    for (const num of descendants(doc, 'num')) {
        const levels = new Map(abstracts.get(attr(child(num, 'abstractNumId'), 'val')) || []);
        for (const override of children(num, 'lvlOverride')) {
            const level = child(override, 'lvl');
            if (level) levels.set(Number(attr(override, 'ilvl')), {
                numFmt: attr(child(level, 'numFmt'), 'val'),
                lvlText: attr(child(level, 'lvlText'), 'val')
            });
        }
        formats.set(attr(num, 'numId'), levels);
    }
    return formats;
}

function runText(run) {
    return Array.from(run.childNodes).map(node => {
        if (node.nodeType !== 1) return '';
        if (node.localName === 't') return node.textContent;
        if (node.localName === 'br' || node.localName === 'cr') return '\n';
        if (node.localName === 'tab') return '\t';
        if (node.localName === 'noBreakHyphen') return '‑';
        return '';
    }).join('');
}

function paragraphRuns(paragraph) {
    // Runs directly in the paragraph or nested in hyperlinks/smart tags/fields.
    return descendants(paragraph, 'r').filter(run => {
        for (let node = run.parentNode; node && node !== paragraph; node = node.parentNode) {
            if (node.localName === 'p') return false;
        }
        return true;
    });
}

/**
 * Builds the accepted-view document model used by every check.
 *
 * @param {Uint8Array} bytes - exported .docx (may contain tracked changes)
 */
export async function buildDocumentModel(bytes) {
    const tracked = unzipDocx(bytes);
    const trackedXml = decoder.decode(tracked.get('word/document.xml'));
    const revisionCount = (trackedXml.match(/<w:(ins|del|rPrChange|pPrChange)\b/g) || []).length;
    const revisionAuthors = {};
    for (const match of trackedXml.matchAll(/<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|numberingChange|tblPrChange|trPrChange|tcPrChange|sectPrChange)\b[^>]*\bw:author="([^"]*)"/g)) {
        revisionAuthors[match[1]] = (revisionAuthors[match[1]] || 0) + 1;
    }

    const acceptedDoc = openDocx(bytes);
    const resolved = await acceptedDoc.resolveRevisions('accept', { allAuthors: true });
    if (resolved.status !== 'ok') throw new Error(`Accept All failed: ${JSON.stringify(resolved.error)}`);
    const entries = unzipDocx(acceptedDoc.toUint8Array());
    const part = name => entries.has(name) ? decoder.decode(entries.get(name)) : null;

    const styles = buildStyles(part('word/styles.xml'));
    const numbering = buildNumbering(part('word/numbering.xml'));
    const doc = parse(part('word/document.xml'), 'word/document.xml');
    const body = descendants(doc, 'body')[0];

    const paragraphs = [];
    const tables = [];
    const visit = (node, tableContext) => {
        for (const element of Array.from(node.childNodes).filter(item => item.nodeType === 1)) {
            if (element.localName === 'p') {
                const pPr = child(element, 'pPr');
                const pStyle = attr(child(pPr, 'pStyle'), 'val');
                const paragraphFormat = styleFormat(styles, pStyle);
                const runs = paragraphRuns(element).map(run => {
                    const rPr = child(run, 'rPr');
                    const format = {
                        ...paragraphFormat,
                        ...styleFormat(styles, attr(child(rPr, 'rStyle'), 'val')),
                        ...declaredFormat(rPr)
                    };
                    return { text: runText(run), format: Object.fromEntries(FORMAT_KEYS.map(key => [key, format[key] === true])) };
                });
                const numPr = child(pPr, 'numPr') || styles.get(pStyle)?.numPr || null;
                const numId = attr(child(numPr, 'numId'), 'val');
                const level = Number(attr(child(numPr, 'ilvl'), 'val') ?? 0);
                const list = numId && numId !== '0' ? {
                    numId,
                    level,
                    numFmt: numbering.get(numId)?.get(level)?.numFmt ?? null
                } : null;
                paragraphs.push({
                    index: paragraphs.length + 1,
                    text: runs.map(run => run.text).join(''),
                    runs,
                    list,
                    table: tableContext
                });
                tableContext?.cell.paragraphs.push(paragraphs.at(-1));
            } else if (element.localName === 'tbl') {
                const table = { index: tables.length, rows: [] };
                tables.push(table);
                children(element, 'tr').forEach((row, rowIndex) => {
                    const rowModel = [];
                    table.rows.push(rowModel);
                    children(row, 'tc').forEach((cell, cellIndex) => {
                        const cellModel = { paragraphs: [] };
                        rowModel.push(cellModel);
                        visit(cell, { table: table.index, row: rowIndex, cell: cellModel, column: cellIndex });
                    });
                });
            } else if (['sdt', 'sdtContent', 'customXml'].includes(element.localName)) {
                visit(element, tableContext);
            }
        }
    };
    visit(body, null);
    for (const table of tables) {
        table.rows = table.rows.map(row => row.map(cell => cell.paragraphs.map(paragraph => paragraph.text).join('\n')));
    }

    const commentsXml = part('word/comments.xml');
    const comments = commentsXml
        ? descendants(parse(commentsXml, 'word/comments.xml'), 'comment').map(comment => ({
            id: attr(comment, 'id'),
            text: descendants(comment, 't').map(node => node.textContent).join('')
        }))
        : [];
    const anchors = descendants(doc, 'commentRangeStart').map(node => attr(node, 'id'));

    return { paragraphs, tables, comments, commentAnchors: anchors, revisionCount, revisionAuthors };
}

/**
 * Revisions added since the previous step must carry the configured redline
 * author. Native Word API edits are stamped with the Office user instead.
 */
export function authorAttribution(previousAuthors, model, configuredAuthor) {
    const added = Object.entries(model.revisionAuthors)
        .map(([author, count]) => [author, count - (previousAuthors?.[author] || 0)])
        .filter(([author, count]) => count > 0 && author !== configuredAuthor);
    return {
        ok: added.length === 0,
        detail: added.length
            ? `new revisions by ${added.map(([author, count]) => `"${author}" ×${count}`).join(', ')} instead of "${configuredAuthor}"`
            : `all new revisions by "${configuredAuthor}"`
    };
}

// ---------------------------------------------------------------------------
// Matching helpers
// ---------------------------------------------------------------------------

const normalize = value => String(value ?? '').replace(/[‘’]/g, "'").replace(/[“”]/g, '"')
    .replace(/\s+/g, ' ').trim();

function textMatcher(spec) {
    if (spec == null) return () => true;
    if (typeof spec === 'string') return text => normalize(text).includes(normalize(spec));
    if (spec.regex) {
        const regex = new RegExp(spec.regex, spec.flags ?? 'i');
        return text => regex.test(normalize(text));
    }
    if (spec.equals != null) return text => normalize(text) === normalize(spec.equals);
    if (spec.startsWith != null) return text => normalize(text).startsWith(normalize(spec.startsWith));
    if (spec.includes != null) return text => normalize(text).includes(normalize(spec.includes));
    throw new Error(`Unsupported text matcher: ${JSON.stringify(spec)}`);
}

const describe = spec => typeof spec === 'string' ? JSON.stringify(spec) : JSON.stringify(spec);
const snippet = text => JSON.stringify(normalize(text).slice(0, 90));

function findParagraphs(model, spec, { inTable } = {}) {
    const matches = textMatcher(spec);
    return model.paragraphs.filter(paragraph => matches(paragraph.text)
        && (inTable === undefined || Boolean(paragraph.table) === inTable));
}

/** Character-level formatting for every occurrence of `text` in a paragraph. */
function occurrenceFormats(paragraph, text) {
    const needle = String(text);
    const characters = paragraph.runs.flatMap(run => Array.from(run.text).map(character => ({ character, format: run.format })));
    const joined = characters.map(item => item.character).join('');
    const results = [];
    for (let at = joined.indexOf(needle); at >= 0; at = joined.indexOf(needle, at + 1)) {
        const slice = characters.slice(at, at + needle.length).filter(item => item.character.trim());
        results.push(Object.fromEntries(FORMAT_KEYS.map(key => [key, slice.length > 0 && slice.every(item => item.format[key])])));
    }
    return results;
}

// ---------------------------------------------------------------------------
// Checks
// ---------------------------------------------------------------------------

const CHECKS = {
    /** A paragraph (optionally at an exact accepted-view index) matches text. */
    paragraph(model, check) {
        if (check.index != null) {
            const paragraph = model.paragraphs[check.index - 1];
            const ok = Boolean(paragraph) && textMatcher(check.text)(paragraph.text);
            return { ok, detail: `P${check.index} is ${paragraph ? snippet(paragraph.text) : 'missing'}` };
        }
        const found = findParagraphs(model, check.text, check);
        const count = check.count ?? null;
        const ok = count == null ? found.length > 0 : found.length === count;
        return { ok, detail: `${found.length} paragraph(s) match ${describe(check.text)}` };
    },

    /** Paragraphs matching `first` and `second` are adjacent, in that order. */
    adjacent(model, check) {
        const first = findParagraphs(model, check.first);
        const ok = first.some(paragraph => {
            const next = model.paragraphs[paragraph.index];
            return next && textMatcher(check.second)(next.text);
        });
        return { ok, detail: `${describe(check.first)} ${ok ? 'is' : 'is not'} immediately followed by ${describe(check.second)}` };
    },

    /**
     * Document-wide text presence/absence. Matched against every line (paragraphs
     * and soft-break lines, so `^` anchors a line start) and the whole document.
     */
    text(model, check) {
        const all = model.paragraphs.map(paragraph => paragraph.text).join('\n');
        const matches = textMatcher(check.text);
        const present = all.split('\n').some(line => matches(line)) || matches(all);
        const ok = check.absent ? !present : present;
        return { ok, detail: `${describe(check.text)} is ${present ? 'present' : 'absent'}` };
    },

    /**
     * Effective character formatting of `find` inside matching paragraphs.
     * `occurrence` (1-based) selects one occurrence; `all: true` requires every
     * occurrence across every matching paragraph.
     */
    format(model, check) {
        const paragraphs = findParagraphs(model, check.paragraph ?? check.find, check);
        const formats = paragraphs.flatMap(paragraph => occurrenceFormats(paragraph, check.find)
            .map((format, index) => ({ format, occurrence: index + 1 })));
        const selected = check.occurrence != null ? formats.filter(item => item.occurrence === check.occurrence) : formats;
        if (selected.length === 0) return { ok: false, detail: `${JSON.stringify(check.find)} not found` };
        const wanted = Object.entries(check.expect);
        const matches = item => wanted.every(([key, value]) => item.format[key] === value);
        const ok = check.all === false || (check.occurrence == null && check.all !== true)
            ? selected.some(matches)
            : selected.every(matches);
        return { ok, detail: `${JSON.stringify(check.find)} formats ${JSON.stringify(selected.map(item => item.format))}, expected ${JSON.stringify(check.expect)}` };
    },

    /**
     * Consecutive real list paragraphs whose texts match `items` in order.
     * Optional numFmt (e.g. upperLetter, decimal, bullet) and level per item.
     */
    list(model, check) {
        const matchers = check.items.map(textMatcher);
        for (let start = 0; start + matchers.length <= model.paragraphs.length; start++) {
            const window = model.paragraphs.slice(start, start + matchers.length);
            if (!window.every((paragraph, offset) => paragraph.list && matchers[offset](paragraph.text))) continue;
            if (check.sameList !== false && new Set(window.map(paragraph => paragraph.list.numId)).size !== 1) continue;
            if (check.levels && !window.every((paragraph, offset) => check.levels[offset] == null || paragraph.list.level === check.levels[offset])) continue;
            if (check.numFmt && !window.every(paragraph => paragraph.list.numFmt === check.numFmt)) continue;
            return { ok: true, detail: `list at P${window[0].index}-P${window.at(-1).index}` };
        }
        const near = model.paragraphs.filter(paragraph => matchers.some(matches => matches(paragraph.text)))
            .map(paragraph => `P${paragraph.index}${paragraph.list ? `[${paragraph.list.numFmt}/L${paragraph.list.level}/#${paragraph.list.numId}]` : '[not a list]'} ${snippet(paragraph.text)}`);
        return { ok: false, detail: `no consecutive list matched; candidates: ${near.join('; ') || 'none'}` };
    },

    /** The paragraph matching `text` is a list item (optionally with level/numFmt). */
    listItem(model, check) {
        const found = findParagraphs(model, check.text);
        const ok = found.some(paragraph => paragraph.list
            && (check.level == null || paragraph.list.level === check.level)
            && (check.numFmt == null || paragraph.list.numFmt === check.numFmt)
            && (check.notNumFmt == null || paragraph.list.numFmt !== check.notNumFmt));
        return { ok, detail: found.map(paragraph => `P${paragraph.index} ${paragraph.list ? JSON.stringify(paragraph.list) : 'not a list'}`).join('; ') || 'not found' };
    },

    /** Every paragraph matching each of `texts` is a list item, all in one list. */
    sameList(model, check) {
        const found = check.texts.map(spec => findParagraphs(model, spec)[0] || null);
        const missing = check.texts.filter((spec, offset) => !found[offset]?.list);
        const ids = new Set(found.filter(paragraph => paragraph?.list).map(paragraph => paragraph.list.numId));
        const formats = new Set(found.filter(paragraph => paragraph?.list).map(paragraph => paragraph.list.numFmt));
        const ok = missing.length === 0 && ids.size === 1
            && (check.numFmt == null || (formats.size === 1 && formats.has(check.numFmt)));
        return { ok, detail: `not list items: ${JSON.stringify(missing)}; numIds ${JSON.stringify([...ids])}; formats ${JSON.stringify([...formats])}` };
    },

    /** Paragraph matching `text` is not a list item (e.g. manual markers removed). */
    notListItem(model, check) {
        const found = findParagraphs(model, check.text);
        const ok = found.length > 0 && found.every(paragraph => !paragraph.list);
        return { ok, detail: found.map(paragraph => `P${paragraph.index} ${paragraph.list ? 'list' : 'plain'}`).join('; ') || 'not found' };
    },

    /**
     * Section content: paragraphs after the header matching `header` up to the
     * next paragraph matching `until`. Requires `minBullets` bullet items and
     * all `mentions` regexes somewhere in the section.
     */
    section(model, check) {
        const header = findParagraphs(model, check.header)[0];
        if (!header) return { ok: false, detail: `header ${describe(check.header)} not found` };
        const stop = textMatcher(check.until);
        const body = [];
        for (const paragraph of model.paragraphs.slice(header.index)) {
            if (stop(paragraph.text)) break;
            if (normalize(paragraph.text)) body.push(paragraph);
        }
        const bullets = body.filter(paragraph => paragraph.list?.numFmt === 'bullet');
        const firstIsIntro = body.length > 0 && !body[0].list;
        const text = body.map(paragraph => paragraph.text).join('\n');
        const missing = (check.mentions || []).filter(pattern => !new RegExp(pattern, 'i').test(text));
        const ok = bullets.length >= (check.minBullets ?? 0)
            && (!check.introParagraph || firstIsIntro)
            && missing.length === 0;
        return { ok, detail: `${body.length} paragraphs, ${bullets.length} bullets, intro=${firstIsIntro}, missing mentions: ${JSON.stringify(missing)}` };
    },

    /** A table row contains cells matching every entry of `row`. */
    tableRow(model, check) {
        const matchers = check.row.map(textMatcher);
        const tables = model.tables.filter(table => check.tableContaining == null
            || table.rows.flat().some(textMatcher(check.tableContaining)));
        const ok = tables.some(table => table.rows.some(row => matchers.every(matches => row.some(cell => matches(cell)))));
        return { ok, detail: `${tables.length} candidate table(s); rows: ${JSON.stringify(tables.map(table => table.rows.map(row => row.map(cell => normalize(cell).slice(0, 30)))))}` };
    },

    /** Table containing `tableContaining` has `rows` rows (or at least `minRows`). */
    tableShape(model, check) {
        const table = model.tables.find(item => item.rows.flat().some(textMatcher(check.tableContaining)));
        if (!table) return { ok: false, detail: `no table contains ${describe(check.tableContaining)}` };
        const ok = (check.rows == null || table.rows.length === check.rows)
            && (check.minRows == null || table.rows.length >= check.minRows);
        return { ok, detail: `table has ${table.rows.length} rows` };
    },

    /** Number of comments, bounded by min/max. */
    comments(model, check) {
        const count = model.comments.length;
        const ok = count >= (check.min ?? 1) && count <= (check.max ?? Infinity) && model.commentAnchors.length >= Math.min(count, check.min ?? 1);
        return { ok, detail: `${count} comment(s), ${model.commentAnchors.length} anchor(s)` };
    },

    /** Some highlighted text matches `text`. */
    highlight(model, check) {
        const matches = textMatcher(check.text);
        const highlighted = model.paragraphs.flatMap(paragraph => {
            const spans = [];
            let current = '';
            for (const run of paragraph.runs) {
                if (run.format.highlight) current += run.text;
                else if (current) { spans.push(current); current = ''; }
            }
            if (current) spans.push(current);
            return spans;
        });
        const ok = highlighted.some(span => matches(span));
        return { ok, detail: `highlighted spans: ${JSON.stringify(highlighted.map(span => normalize(span).slice(0, 60)))}` };
    }
};

/** Evaluate one declarative check against a model. */
export function evaluateCheck(model, check) {
    const evaluate = CHECKS[check.type];
    if (!evaluate) return { ok: false, detail: `unknown check type ${check.type}` };
    try {
        return evaluate(model, check);
    } catch (error) {
        return { ok: false, detail: `check threw: ${error.message}` };
    }
}

/**
 * Score one step: its own checks plus regression of earlier persistent checks.
 *
 * @param {object} model - buildDocumentModel() result for the document after `stepIndex`
 * @param {Array} steps - scenario steps
 * @param {number} stepIndex - 0-based index of the step just completed
 * @param {Set<string>|null} [passedBefore] - keys of checks that passed at their own
 *   step; earlier checks are regressed only when they once passed (a failing
 *   step is reported once, not repeated at every later step)
 */
export function scoreStep(model, steps, stepIndex, passedBefore = null) {
    const results = [];
    steps.slice(0, stepIndex + 1).forEach((step, index) => {
        step.checks.forEach((check, checkIndex) => {
            const own = index === stepIndex;
            const key = checkKey(step, checkIndex);
            if (!own && (check.persist === false || step.persist === false)) return;
            if (!own && passedBefore && !passedBefore.has(key)) return;
            const outcome = evaluateCheck(model, check);
            results.push({
                key,
                step: step.id,
                own,
                label: check.label || check.type,
                ok: outcome.ok,
                detail: outcome.detail,
                ...(step.knownIssue ? { knownIssue: step.knownIssue } : {})
            });
        });
    });
    return results;
}

export const checkKey = (step, checkIndex) => `${step.id}#${checkIndex}`;
