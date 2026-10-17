/** Word tracking-state adapter shared by the checkpointed host writer. */

/**
 * Runs one host operation with native change tracking temporarily disabled.
 * Once the disable assignment is queued, restoration is attempted even if its
 * synchronization fails. A restoration error never hides an earlier failure.
 *
 * @template T
 * @param {Word.RequestContext} context
 * @param {(state: { originalMode: Word.ChangeTrackingMode, trackingDisabled: boolean }) => Promise<T>} operation
 * @param {{ enabled?: boolean, logPrefix?: string }} [options]
 * @returns {Promise<T>}
 */
export async function withNativeTrackingDisabled(context, operation, options = {}) {
    const {
        enabled = true,
        logPrefix = 'Tracking'
    } = options;

    context.document.load('changeTrackingMode');
    await context.sync();
    const originalMode = context.document.changeTrackingMode;

    const trackingDisabled = !!enabled && originalMode !== Word.ChangeTrackingMode.off;
    let primaryError = null;
    let value;

    if (trackingDisabled) {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        try {
            await context.sync();
        } catch (error) {
            primaryError = error;
        }
    }

    if (!primaryError) {
        try {
            value = await operation({ originalMode, trackingDisabled });
        } catch (error) {
            primaryError = error;
        }
    }

    let restoreError = null;
    if (trackingDisabled) {
        try {
            context.document.changeTrackingMode = originalMode;
            await context.sync();
        } catch (error) {
            restoreError = error;
        }
    } else if (enabled) {
        console.log(`[${logPrefix}] Native tracking already off; no toggle needed`);
    }

    if (primaryError) {
        if (restoreError && primaryError && typeof primaryError === 'object') {
            try { primaryError.trackingRestoreError = restoreError; } catch { /* Preserve the original error. */ }
        }
        throw primaryError;
    }
    if (restoreError) throw restoreError;
    return value;
}
