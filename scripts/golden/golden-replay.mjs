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

function currentStepMessages(payload, requestPrompt) {
    const messages = Array.isArray(payload?.contents) ? payload.contents : [];
    if (typeof requestPrompt !== 'string' || !requestPrompt) return messages;
    let requestIndex = -1;
    for (let messageIndex = 0; messageIndex < messages.length; messageIndex++) {
        const message = messages[messageIndex];
        if (message?.role !== 'user') continue;
        const text = (Array.isArray(message.parts) ? message.parts : [])
            .filter(part => typeof part?.text === 'string')
            .map(part => part.text)
            .join('\n');
        if (text.includes(requestPrompt)) requestIndex = messageIndex;
    }
    return requestIndex < 0 ? [] : messages.slice(requestIndex);
}

function latestFunctionResponse(payload, toolName, requestPrompt = null) {
    const messages = currentStepMessages(payload, requestPrompt);
    for (let messageIndex = messages.length - 1; messageIndex >= 0; messageIndex--) {
        const parts = Array.isArray(messages[messageIndex]?.parts) ? messages[messageIndex].parts : [];
        for (let partIndex = parts.length - 1; partIndex >= 0; partIndex--) {
            const functionResponse = parts[partIndex]?.functionResponse;
            if (functionResponse?.name === toolName) return functionResponse;
        }
    }
    return null;
}

function parseFunctionResponse(functionResponse) {
    const content = functionResponse?.response?.content;
    const text = Array.isArray(content)
        ? content.map(part => typeof part?.text === 'string' ? part.text : '').join('\n').trim()
        : '';
    if (!text) return null;
    try { return JSON.parse(text); } catch { return text; }
}

/** Parse the latest function result after this request's prompt, ignoring prior request history. */
export function extractLatestFunctionResponse(payload, toolName, requestPrompt = null) {
    return parseFunctionResponse(latestFunctionResponse(payload, toolName, requestPrompt));
}

/** All parsed function results after a request's own prompt, in history order. */
export function extractFunctionResponses(payload, toolName, requestPrompt = null) {
    const results = [];
    for (const message of currentStepMessages(payload, requestPrompt)) {
        for (const part of Array.isArray(message?.parts) ? message.parts : []) {
            if (part?.functionResponse?.name === toolName) results.push(parseFunctionResponse(part.functionResponse));
        }
    }
    return results;
}

/** The latest inspection result actually present in Gemini's function-response history. */
export function extractLatestInspection(payload, requestPrompt = null) {
    const response = extractLatestFunctionResponse(payload, 'inspect_document', requestPrompt);
    return response && typeof response === 'object' ? response : null;
}

function inspectionForHandle(payload, requestPrompt) {
    const functionResponse = latestFunctionResponse(payload, 'inspect_document', requestPrompt);
    if (!functionResponse) throw new Error('$handle/$end requires an inspect_document response in the current model history');
    const inspection = parseFunctionResponse(functionResponse);
    if (!inspection || inspection.status !== 'ok') {
        throw new Error('The latest inspect_document response did not contain a successful JSON result');
    }
    return inspection;
}

function resolveHandle(inspection, spec, nth = 1) {
    const matches = (Array.isArray(inspection.targets) ? inspection.targets : [])
        .filter(target => typeof target?.handle === 'string' && matcher(spec)(target.text));
    const target = matches[nth - 1];
    if (!target) {
        throw new Error(`Replay handle reference ${JSON.stringify(spec)}${nth > 1 ? ` (#${nth})` : ''} matched ${matches.length} inspected target(s)`);
    }
    return target.handle;
}

function resolveDocumentEnd(inspection) {
    const handle = inspection.documentEnd?.handle;
    if (typeof handle !== 'string' || !handle) throw new Error('Replay $end reference requires an inspected document-end handle');
    return handle;
}

/** Deep-resolve legacy { $p }, { $all }, { $anchor } and direct { $handle }/{ $end } references. */
export function resolveReferences(value, context, payload = null, requestPrompt = null) {
    if (Array.isArray(value)) return value.map(item => resolveReferences(item, context, payload, requestPrompt));
    if (!value || typeof value !== 'object') return value;
    if ('$p' in value) return resolveParagraph(context, value.$p, value.nth).index;
    if ('$handle' in value) return resolveHandle(inspectionForHandle(payload, requestPrompt), value.$handle, value.nth);
    if ('$end' in value) return resolveDocumentEnd(inspectionForHandle(payload, requestPrompt));
    // Copy the anchor verbatim (whitespace collapsed only), as the prompt instructs.
    if ('$anchor' in value) return resolveParagraph(context, value.$anchor, value.nth).text.replace(/\s+/g, ' ').trim().slice(0, ANCHOR_LENGTH);
    if ('$all' in value) {
        const indices = context.filter(paragraph => matcher(value.$all)(paragraph.text)).map(paragraph => paragraph.index);
        if (indices.length === 0) throw new Error(`Replay reference $all ${JSON.stringify(value.$all)} matched no paragraphs`);
        return indices;
    }
    return Object.fromEntries(Object.entries(value).map(([key, nested]) => [key, resolveReferences(nested, context, payload, requestPrompt)]));
}

/** Wrap a scripted entry as a Gemini generateContent response body. */
export function buildResponse(entry, context, payload = null, requestPrompt = null) {
    const resolved = resolveReferences(entry.chat ?? entry.aux, context, payload, requestPrompt);
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
    const requestCounts = { chat: 0, aux: 0, other: 0 };
    let functionResponses = {};
    let functionResponseHistory = {};
    let lastContext = [];
    return {
        next(payload) {
            const kind = classifyRequest(payload);
            requestCounts[kind] = (requestCounts[kind] || 0) + 1;
            functionResponses = {};
            functionResponseHistory = {};
            for (const message of currentStepMessages(payload, step.prompt)) {
                for (const part of Array.isArray(message?.parts) ? message.parts : []) {
                    const functionResponse = part?.functionResponse;
                    if (typeof functionResponse?.name === 'string') {
                        const name = functionResponse.name;
                        const result = parseFunctionResponse(functionResponse);
                        functionResponses[name] = result;
                        (functionResponseHistory[name] ||= []).push(result);
                    }
                }
            }
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
                const response = buildResponse(entry, context.length ? context : lastContext, payload, step.prompt);
                served.push({ kind, part: response.candidates[0].content.parts[0] });
                return { kind, entry, response };
            } catch (error) {
                desyncs.push(error.message);
                const fallback = kind === 'chat' ? { chat: { text: 'Done.' } } : { aux: [] };
                return { kind, error: error.message, response: buildResponse(fallback, lastContext) };
            }
        },
        summary() {
            return { remaining: queue.length, desyncs, served, requestCounts, functionResponses, functionResponseHistory };
        }
    };
}
