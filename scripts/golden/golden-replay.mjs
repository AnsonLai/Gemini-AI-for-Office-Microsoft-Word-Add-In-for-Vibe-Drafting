// Deterministic model replay for the golden scenario. Classifies intercepted
// Gemini generateContent payloads, resolves runtime paragraph references from the
// [P#] context inside each request, and wraps scripted parts as Gemini responses.

const ANCHOR_LENGTH = 40;

/** 'chat' (function-calling chat model), 'aux' (tool JSON calls), or 'other' (Glance/research). */
export function classifyRequest(payload) {
    const tools = Array.isArray(payload?.tools) ? payload.tools : [];
    if (tools.some(tool => Array.isArray(tool?.functionDeclarations) || Array.isArray(tool?.function_declarations))) return 'chat';
    if (payload?.generationConfig?.responseSchema || payload?.generationConfig?.responseMimeType === 'application/json') return 'aux';
    return 'other';
}

function textParts(payload) {
    const parts = [];
    const visit = value => {
        if (!value || typeof value !== 'object') return;
        if (typeof value.text === 'string') parts.push(value.text);
        for (const nested of Object.values(value)) if (nested && typeof nested === 'object') visit(nested);
    };
    visit(payload?.systemInstruction);
    visit(payload?.contents);
    return parts;
}

/**
 * The newest [P#] document view in a request: paragraphs as { index, meta, text }.
 * Continuation lines (soft breaks rendered as newlines) join their paragraph.
 */
export function extractContext(payload) {
    // Every document view (chat context, tool prompts, refreshed context after
    // an edit) is embedded as a """...""" block; the last one is the newest.
    const blocks = textParts(payload).flatMap(text => [...text.matchAll(/"""([\s\S]*?)"""/g)].map(match => match[1]))
        .filter(block => /^\s*\[P1(\|[^\]]*)?\]/.test(block));
    const latest = blocks.at(-1);
    if (!latest) return [];
    const paragraphs = [];
    for (const line of latest.replace(/^\s+/, '').split(/\r?\n/)) {
        const match = line.match(/^\[P(\d+)(?:\|([^\]]*))?\]\s?(.*)$/);
        if (match) paragraphs.push({ index: Number(match[1]), meta: match[2] || '', text: match[3] });
        else if (paragraphs.length) paragraphs.at(-1).text += `\n${line}`;
    }
    return paragraphs;
}

const normalize = value => String(value ?? '').replace(/[‘’]/g, "'").replace(/[“”]/g, '"')
    .replace(/\s+/g, ' ').trim();

function matcher(spec) {
    if (typeof spec === 'string') return text => normalize(text).includes(normalize(spec));
    if (spec.regex) {
        const regex = new RegExp(spec.regex, spec.flags ?? 'i');
        return text => regex.test(normalize(text));
    }
    if (spec.equals != null) return text => normalize(text) === normalize(spec.equals);
    if (spec.startsWith != null) return text => normalize(text).startsWith(normalize(spec.startsWith));
    if (spec.includes != null) return text => normalize(text).includes(normalize(spec.includes));
    throw new Error(`Unsupported replay matcher: ${JSON.stringify(spec)}`);
}

function resolveParagraph(context, spec, nth = 1) {
    const matches = context.filter(paragraph => matcher(spec)(paragraph.text));
    const paragraph = matches[nth - 1];
    if (!paragraph) {
        throw new Error(`Replay reference ${JSON.stringify(spec)}${nth > 1 ? ` (#${nth})` : ''} matched ${matches.length} paragraph(s) in the request context`);
    }
    return paragraph;
}

/** Deep-resolve { $p }, { $all }, { $anchor } placeholders against a context. */
export function resolveReferences(value, context) {
    if (Array.isArray(value)) return value.map(item => resolveReferences(item, context));
    if (!value || typeof value !== 'object') return value;
    if ('$p' in value) return resolveParagraph(context, value.$p, value.nth).index;
    // Copy the anchor verbatim (whitespace collapsed only), as the prompt instructs.
    if ('$anchor' in value) return resolveParagraph(context, value.$anchor, value.nth).text.replace(/\s+/g, ' ').trim().slice(0, ANCHOR_LENGTH);
    if ('$all' in value) {
        const indices = context.filter(paragraph => matcher(value.$all)(paragraph.text)).map(paragraph => paragraph.index);
        if (indices.length === 0) throw new Error(`Replay reference $all ${JSON.stringify(value.$all)} matched no paragraphs`);
        return indices;
    }
    return Object.fromEntries(Object.entries(value).map(([key, nested]) => [key, resolveReferences(nested, context)]));
}

/** Wrap a scripted entry as a Gemini generateContent response body. */
export function buildResponse(entry, context) {
    const resolved = resolveReferences(entry.chat ?? entry.aux, context);
    const part = entry.chat != null ? resolved : { text: JSON.stringify(resolved) };
    return { candidates: [{ content: { role: 'model', parts: [part] }, finishReason: 'STOP', index: 0 }] };
}

/**
 * Per-step replay queue. `next(payload)` returns { response, entry, kind, error? }.
 * A kind mismatch or exhausted queue is a desync: it is recorded and answered
 * with a harmless reply so the chat loop can finish and the step is scored.
 */
export function createReplayQueue(step) {
    const queue = [...(step.replay || [])];
    const desyncs = [];
    const served = [];
    let lastContext = [];
    return {
        next(payload) {
            const kind = classifyRequest(payload);
            const context = extractContext(payload);
            if (context.length) lastContext = context;
            if (kind === 'other') {
                return { kind, response: { candidates: [{ content: { role: 'model', parts: [{ text: '{}' }] }, finishReason: 'STOP', index: 0 }] } };
            }
            const entry = queue[0];
            const entryKind = entry ? (entry.chat != null ? 'chat' : 'aux') : null;
            if (!entry || entryKind !== kind) {
                const error = !entry ? `unscripted ${kind} call (replay exhausted)` : `expected ${entryKind} call, got ${kind}`;
                desyncs.push(error);
                const fallback = kind === 'chat' ? { chat: { text: 'Done.' } } : { aux: [] };
                return { kind, error, response: buildResponse(fallback, lastContext) };
            }
            queue.shift();
            try {
                const response = buildResponse(entry, context.length ? context : lastContext);
                served.push({ kind, part: response.candidates[0].content.parts[0] });
                return { kind, entry, response };
            } catch (error) {
                desyncs.push(error.message);
                const fallback = kind === 'chat' ? { chat: { text: 'Done.' } } : { aux: [] };
                return { kind, error: error.message, response: buildResponse(fallback, lastContext) };
            }
        },
        summary() {
            return { remaining: queue.length, desyncs, served };
        }
    };
}
