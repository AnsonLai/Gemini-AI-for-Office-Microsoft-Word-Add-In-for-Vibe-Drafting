/** Word host adapter for one body-level canonical OOXML batch. */
import { prepareWordFlatOpcBatch } from '@ansonlai/docx-redline-js/word';
import { withNativeTrackingDisabled } from './word-ooxml.js';

export { inspectWordFlatOpc, prepareWordFlatOpcBatch } from '@ansonlai/docx-redline-js/word';

function failedBatchWrite(result, source, code, error, writeAttempted = false, written = false) {
    return {
        ...result,
        status: 'error',
        written,
        writeAttempted,
        mutationOutcome: written ? 'applied_with_host_error' : writeAttempted ? 'indeterminate' : 'prepared',
        engineResult: result,
        error: { code, message: error?.message || String(error) },
        hostError: {
            ...(error?.code ? { code: error.code } : {}),
            message: error?.message || String(error)
        },
        source
    };
}

function createBatchTiming(options) {
    if (typeof options.onTiming !== 'function') return null;
    const observer = options.onTiming;
    const clock = typeof options.now === 'function'
        ? options.now
        : () => globalThis.performance?.now?.();
    const now = () => {
        try {
            const value = clock();
            return Number.isFinite(value) ? value : null;
        } catch {
            return null;
        }
    };
    return {
        now,
        record(phase, startedAt) {
            if (startedAt === undefined || startedAt === null) return;
            const endedAt = now();
            if (endedAt === null) return;
            const durationMs = Math.max(0, endedAt - startedAt);
            try { observer({ phase, durationMs }); } catch { /* Diagnostics must not affect Word edits. */ }
        }
    };
}

/**
 * Reads a Word body once, prepares a canonical operation batch, and performs at
 * most one whole-body package replacement.
 *
 * `operations` may be an array or an async factory receiving the inspected
 * source. Partial ranges and paragraphs are intentionally unsupported.
 */
export async function executePureOoxmlBatch(context, body, operations, options = {}) {
    const timing = createBatchTiming(options);
    const startedAt = timing?.now();
    try {
        return await executePureOoxmlBatchInternal(context, body, operations, options, timing);
    } finally {
        timing?.record('adapterTotal', startedAt);
    }
}

async function executePureOoxmlBatchInternal(context, body, operations, options, timing) {
    if (!context || typeof context.sync !== 'function') {
        throw new TypeError('A Word request context is required');
    }
    if (!body || typeof body.getOoxml !== 'function' || typeof body.insertOoxml !== 'function') {
        throw new TypeError('A Word body is required for a whole-document batch');
    }
    const checkCancelled = () => {
        if (options.signal?.aborted) throw Object.assign(new Error('Request cancelled before document commit.'), { name: 'AbortError', code: 'REQUEST_CANCELLED' });
    };
    checkCancelled();
    let ooxmlResult;
    const sourceStartedAt = timing?.now();
    try {
        ooxmlResult = body.getOoxml();
        await context.sync();
    } finally {
        timing?.record('sourceReadSync', sourceStartedAt);
    }
    let preparation;
    const preparationStartedAt = timing?.now();
    try {
        const preparationOptions = { ...options };
        delete preparationOptions.onTiming;
        delete preparationOptions.now;
        preparation = await prepareWordFlatOpcBatch(ooxmlResult?.value || '', operations, preparationOptions);
    } finally {
        timing?.record('portablePreparation', preparationStartedAt);
    }
    const publicSource = preparation.source;
    const result = preparation.result;
    checkCancelled();
    if (typeof options.onPrepared === 'function') await options.onPrepared(preparation);
    if (preparation.status === 'noop') {
        if (preparation.operations.length === 0) {
            return { status: 'ok', hasChanges: false, written: false, writeAttempted: false, mutationOutcome: 'noop', results: [], receipts: [], source: publicSource };
        }
        return { ...result, written: false, writeAttempted: false, mutationOutcome: 'noop', source: publicSource };
    }
    if (preparation.status === 'refused') {
        return { ...result, written: false, writeAttempted: false, mutationOutcome: result?.rolledBack ? 'rolled_back' : 'refused', source: publicSource };
    }
    if (preparation.status === 'error') {
        return failedBatchWrite(result, publicSource, 'WORD_OOXML_PACKAGE_FAILED', preparation.error);
    }
    let checkpointIndex;
    try {
        if (typeof options.beforeWrite === 'function') {
            checkpointIndex = await options.beforeWrite({ ...preparation, sourceOoxml: ooxmlResult.value });
        }
        checkCancelled();
    } catch (error) {
        return { ...failedBatchWrite(result, publicSource, error.code || 'PRECOMMIT_FAILED', error), mutationOutcome: 'refused', checkpointIndex };
    }
    const insertionPayload = preparation.insertionPayload;
    let writeAttempted = false;
    let written = false;
    const writeBatch = async () => {
        checkCancelled();
        const insertionStartedAt = timing?.now();
        try {
            writeAttempted = true;
            const insertMode = (typeof Word !== 'undefined' && (Word.InsertLocation?.replace || Word.InsertLocation?.Replace))
                || 'Replace';
            body.insertOoxml(insertionPayload, insertMode);
            await context.sync();
            written = true;
        } finally {
            timing?.record('insertSync', insertionStartedAt);
        }
    };
    try {
        if (options.disableNativeTracking) {
            await withNativeTrackingDisabled(context, writeBatch, {
                enabled: true,
                logPrefix: options.logPrefix || 'WordOp/Batch'
            });
        } else {
            await writeBatch();
        }
    } catch (error) {
        return { ...failedBatchWrite(result, publicSource, 'WORD_OOXML_WRITE_FAILED', error, writeAttempted, written), checkpointIndex,
            requiresInspection: options.disableNativeTracking === true && error.code !== 'REQUEST_CANCELLED' };
    }
    return { ...result, written: true, writeAttempted: true, mutationOutcome: 'applied', source: publicSource, checkpointIndex };
}
