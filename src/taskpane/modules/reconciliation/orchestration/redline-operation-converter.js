/**
 * Converts add-in redline tool payloads into shared standalone redline operations.
 */

function normalizeEscapes(content) {
    if (content == null) return '';
    return String(content)
        .replace(/\\n/g, '\n')
        .replace(/\\t/g, '\t')
        .replace(/\\r/g, '\r');
}

function toNonNegativeInteger(value, fallback = 0) {
    const parsed = Number.parseInt(String(value ?? ''), 10);
    return Number.isInteger(parsed) && parsed >= 0 ? parsed : fallback;
}

/**
 * Applies one substring replacement against source text.
 *
 * @param {string} sourceText
 * @param {string} searchText
 * @param {string} replacementText
 * @returns {{ applied: boolean, modifiedText: string, matchMode: 'exact'|'case_insensitive'|null }}
 */
export function applySubstringSearchReplace(sourceText, searchText, replacementText) {
    const source = String(sourceText || '');
    const search = String(searchText || '');
    const replacement = String(replacementText || '');
    if (!search) {
        return {
            applied: false,
            modifiedText: source,
            matchMode: null
        };
    }

    const exactIndex = source.indexOf(search);
    if (exactIndex >= 0) {
        return {
            applied: true,
            modifiedText: `${source.slice(0, exactIndex)}${replacement}${source.slice(exactIndex + search.length)}`,
            matchMode: 'exact'
        };
    }

    const lowerSource = source.toLowerCase();
    const lowerSearch = search.toLowerCase();
    const insensitiveIndex = lowerSource.indexOf(lowerSearch);
    if (insensitiveIndex >= 0) {
        return {
            applied: true,
            modifiedText: `${source.slice(0, insensitiveIndex)}${replacement}${source.slice(insensitiveIndex + search.length)}`,
            matchMode: 'case_insensitive'
        };
    }

    return {
        applied: false,
        modifiedText: source,
        matchMode: null
    };
}

/**
 * Builds a shared redline operation for a scoped paragraph/range OOXML document.
 *
 * @param {Object} change - Add-in tool payload
 * @param {{
 *   scopeStartText?: string,
 *   scopeParagraphCount?: number,
 *   insertionBeforeStart?: boolean
 * }} [context={}]
 * @returns {{ ok: true, operation: Object } | { ok: false, reason: string }}
 */
export function toScopedSharedRedlineOperation(change, context = {}) {
    const operationName = String(change?.operation || '').trim().toLowerCase();
    const scopeStartText = String(context.scopeStartText || '').trim();
    const scopeParagraphCount = Math.max(1, toNonNegativeInteger(context.scopeParagraphCount, 1));
    const insertionBeforeStart = context.insertionBeforeStart === true;

    if (!scopeStartText) {
        return { ok: false, reason: 'Target paragraph text is empty; cannot build shared redline target.' };
    }

    if (!operationName) {
        return { ok: false, reason: 'Missing change.operation value.' };
    }

    let modifiedText = null;
    if (operationName === 'edit_paragraph') {
        if (change?.newContent == null) {
            return { ok: false, reason: 'Missing newContent for edit_paragraph.' };
        }
        modifiedText = normalizeEscapes(change.newContent);
    } else if (operationName === 'replace_paragraph' || operationName === 'replace_range') {
        const replacementContent = change?.content ?? change?.newContent ?? change?.replacementText;
        if (replacementContent == null) {
            return { ok: false, reason: `Missing content for ${operationName}.` };
        }
        const normalizedReplacement = normalizeEscapes(replacementContent);
        if (operationName === 'replace_range' && insertionBeforeStart) {
            // Model occasionally emits an insertion-at-start shape as:
            // replace_range(P1..P0). Normalize to "insert paragraph before P1"
            // by preserving P1 text and prefixing inserted content + paragraph break.
            modifiedText = scopeStartText
                ? `${normalizedReplacement}\n${scopeStartText}`
                : normalizedReplacement;
        } else {
            modifiedText = normalizedReplacement;
        }
    } else if (operationName === 'modify_text') {
        const searchText = normalizeEscapes(change?.originalText || '');
        if (!searchText.trim()) {
            return { ok: false, reason: 'Missing originalText for modify_text.' };
        }
        const replacementText = normalizeEscapes(change?.replacementText || '');
        const replacement = applySubstringSearchReplace(scopeStartText, searchText, replacementText);
        if (!replacement.applied) {
            return { ok: false, reason: `Could not find modify_text originalText in target paragraph.` };
        }
        modifiedText = replacement.modifiedText;
    } else {
        return { ok: false, reason: `Unsupported redline operation: ${operationName}` };
    }

    const op = {
        type: 'redline',
        // Scoped bridge payloads are rebuilt as P1..Pn; always anchor to P1.
        targetRef: 'P1',
        target: scopeStartText,
        modified: modifiedText
    };
    if (operationName === 'replace_range') {
        if (scopeParagraphCount > 1) {
            op.targetEndRef = `P${scopeParagraphCount}`;
        }
    }

    return { ok: true, operation: op };
}
