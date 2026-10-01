/** Word host adapter. OOXML preparation lives in the host-neutral consumer core. */
import { applySharedOperationToParagraphOoxml, applySharedOperationToScopeOoxml, isSimplePlainTextRedline, prepareCanonicalBatch } from './consumer-core.js';
import { insertOoxmlWithRangeFallback, withNativeTrackingDisabled } from './word-ooxml.js';

export {
    captureSourceBaseline,
    captureWordSourceBaseline,
    isSimplePlainTextRedline,
    applySharedOperationToParagraphOoxml,
    applySharedOperationToScopeOoxml,
    prepareCanonicalBatch
} from './consumer-core.js';

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

function resolveWordOperationScope(scope) {
    if (!scope) throw new Error('Missing scope for applyWordOperation');
    if (scope.paragraph) {
        if (scope.endParagraph) {
            const range = scope.paragraph.getRange().expandTo(scope.endParagraph.getRange());
            return { kind: 'range', target: range };
        }
        return { kind: 'paragraph', target: scope.paragraph };
    }
    if (scope.range) {
        return { kind: 'range', target: scope.range };
    }
    if (typeof scope.getOoxml === 'function' && typeof scope.insertOoxml === 'function') {
        return { kind: 'range', target: scope };
    }
    throw new Error('Unsupported scope shape for applyWordOperation');
}

/**
 * Applies one atomic operation batch to a Word body, range, or paragraph.
 * `operations` may be an array or an async factory receiving the one-read source.
 * The engine result and its receipts are returned unchanged, with `written` added.
 */
export async function executePureOoxmlBatch(context, targetScope, operations, options = {}) {
    const timing = createBatchTiming(options);
    const startedAt = timing?.now();
    try {
        return await executePureOoxmlBatchInternal(context, targetScope, operations, options, timing);
    } finally {
        timing?.record('adapterTotal', startedAt);
    }
}

async function executePureOoxmlBatchInternal(context, targetScope, operations, options, timing) {
    if (!context || typeof context.sync !== 'function') {
        throw new TypeError('A Word request context is required');
    }
    const resolved = resolveWordOperationScope(targetScope);
    let ooxmlResult;
    const sourceStartedAt = timing?.now();
    try {
        ooxmlResult = resolved.target.getOoxml();
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
        preparation = await prepareCanonicalBatch(ooxmlResult?.value || '', operations, preparationOptions);
    } finally {
        timing?.record('portablePreparation', preparationStartedAt);
    }
    const publicSource = preparation.source;
    const result = preparation.result;
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
    const insertionPayload = preparation.insertionPayload;
    let writeAttempted = false;
    let written = false;
    const writeBatch = async () => {
        const insertionStartedAt = timing?.now();
        try {
            writeAttempted = true;
            if (resolved.kind === 'paragraph') {
                await insertOoxmlWithRangeFallback(
                    resolved.target,
                    insertionPayload,
                    'Replace',
                    context,
                    options.logPrefix || 'WordOp/Batch'
                );
            } else {
                const insertMode = (typeof Word !== 'undefined' && (Word.InsertLocation?.replace || Word.InsertLocation?.Replace))
                    || 'Replace';
                resolved.target.insertOoxml(insertionPayload, insertMode);
                await context.sync();
            }
            written = true;
        } finally {
            timing?.record('insertSync', insertionStartedAt);
        }
    };
    try {
        if (options.disableNativeTracking) {
            await withNativeTrackingDisabled(context, writeBatch, {
                enabled: true,
                baseTrackingMode: options.baseTrackingMode ?? null,
                logPrefix: options.logPrefix || 'WordOp/Batch'
            });
        } else {
            await writeBatch();
        }
    } catch (error) {
        return failedBatchWrite(result, publicSource, 'WORD_OOXML_WRITE_FAILED', error, writeAttempted, written);
    }
    return { ...result, written: true, writeAttempted: true, mutationOutcome: 'applied', source: publicSource };
}

/**
 * Applies a canonical operation to a Word paragraph/range scope.
 *
 * @param {Word.RequestContext} context
 * @param {Object} operation
 * @param {Object} scope
 * @param {Object} [options={}]
 * @returns {Promise<boolean>} True when changes were applied
 */
export async function applyWordOperation(context, operation, scope, options = {}) {
    const resolved = resolveWordOperationScope(scope);
    const scopeOoxmlResult = resolved.target.getOoxml();
    await context.sync();

    const bridgeOptions = {
        author: options.author,
        generateRedlines: options.generateRedlines,
        sanitizeInput: options.sanitizeInput === true,
        existingRevisions: options.existingRevisions,
        onInfo: options.onInfo,
        onWarn: options.onWarn,
        runner: options.runner
    };
    const bridgeResult = resolved.kind === 'paragraph'
        ? await applySharedOperationToParagraphOoxml(scopeOoxmlResult?.value || '', operation, bridgeOptions)
        : await applySharedOperationToScopeOoxml(scopeOoxmlResult?.value || '', operation, bridgeOptions);

    if (!bridgeResult.hasChanges) {
        return false;
    }

    const canUseDirectParagraphPayload = false; // direct_paragraph snippet injection crashes Word parser with w:ins elements
    const insertionPayload = canUseDirectParagraphPayload
        ? bridgeResult.paragraphOoxml
        : bridgeResult.packageOoxml;
    if (!insertionPayload) {
        return false;
    }
    if (typeof options.onInfo === 'function') {
        options.onInfo(
            `Insertion strategy: ${canUseDirectParagraphPayload ? 'direct_paragraph' : 'package'} `
            + `(kind=${resolved.kind}, singleParagraphOutput=${bridgeResult.singleParagraphOutput === true}, `
            + `hasComments=${!!bridgeResult.commentsXml}, hasNumbering=${!!bridgeResult.numberingXml}, `
            + `isSimplePlainTextRedline=${isSimplePlainTextRedline(operation)})`
        );
    }

    await withNativeTrackingDisabled(context, async () => {
        if (resolved.kind === 'paragraph') {
            await insertOoxmlWithRangeFallback(
                resolved.target,
                insertionPayload,
                'Replace',
                context,
                options.logPrefix || 'WordOp/Shared'
            );
        } else {
            const insertMode = (typeof Word !== 'undefined' && (Word.InsertLocation?.replace || Word.InsertLocation?.Replace))
                || 'Replace';
            resolved.target.insertOoxml(insertionPayload, insertMode);
            await context.sync();
        }
    }, {
        enabled: !!options.disableNativeTracking,
        baseTrackingMode: options.baseTrackingMode ?? null,
        logPrefix: options.logPrefix || 'WordOp/Shared'
    });

    return true;
}

/**
 * Applies a shared standalone operation to a single Word paragraph.
 *
 * @param {Object} params
 * @param {Word.RequestContext} params.context
 * @param {Word.Paragraph} params.targetParagraph
 * @param {Object} params.operation
 * @param {string} params.author
 * @param {boolean} params.generateRedlines
 * @param {boolean} [params.disableNativeTracking=false]
 * @param {Word.ChangeTrackingMode|null} [params.baseTrackingMode=null]
 * @param {string} params.logPrefix
 * @returns {Promise<boolean>} True when a change is applied
 */
export async function applySharedOperationToWordParagraph({
    context,
    targetParagraph,
    operation,
    author,
    generateRedlines,
    disableNativeTracking = false,
    baseTrackingMode = null,
    logPrefix
}) {
    return applyWordOperation(context, operation, { paragraph: targetParagraph }, {
        author,
        generateRedlines,
        disableNativeTracking,
        baseTrackingMode,
        logPrefix,
        onInfo: () => {},
        onWarn: () => console.warn(`[${logPrefix}] Operation warning; consult the structured result.`)
    });
}

/**
 * Applies a shared standalone operation to a Word paragraph/range scope.
 *
 * @param {Object} params
 * @param {Word.RequestContext} params.context
 * @param {Word.Range|Word.Paragraph} params.scope
 * @param {Object} params.operation
 * @param {string} params.author
 * @param {boolean} params.generateRedlines
 * @param {boolean} [params.disableNativeTracking=false]
 * @param {Word.ChangeTrackingMode|null} [params.baseTrackingMode=null]
 * @param {string} params.logPrefix
 * @returns {Promise<boolean>} True when a change is applied
 */
export async function applySharedOperationToWordScope({
    context,
    scope,
    operation,
    author,
    generateRedlines,
    disableNativeTracking = false,
    baseTrackingMode = null,
    logPrefix
}) {
    return applyWordOperation(context, operation, { range: scope }, {
        author,
        generateRedlines,
        disableNativeTracking,
        baseTrackingMode,
        logPrefix,
        onInfo: () => {},
        onWarn: () => console.warn(`[${logPrefix}] Operation warning; consult the structured result.`)
    });
}
