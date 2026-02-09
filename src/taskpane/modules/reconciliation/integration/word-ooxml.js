/**
 * Word-specific OOXML interop helpers.
 *
 * These helpers intentionally depend on Word JS runtime objects.
 */

/**
 * Reads OOXML for a paragraph with fallback chain:
 * paragraph -> paragraph range -> parent cell -> parent cell range -> parent table -> parent table range.
 *
 * @param {Word.Paragraph} paragraph - Target paragraph proxy
 * @param {Word.RequestContext} context - Word request context
 * @param {Object} [options={}] - Read options
 * @param {boolean} [options.propertiesPreloaded=false] - Whether parent table/cell properties are already loaded
 * @param {string} [options.logPrefix='OxmlRead'] - Log prefix
 * @returns {Promise<{ ooxml: string|null, source: string|null }>}
 */
export async function getParagraphOoxmlWithFallback(paragraph, context, options = {}) {
    const {
        propertiesPreloaded = false,
        logPrefix = 'OxmlRead'
    } = options;

    let ooxmlResult = null;
    let ooxmlValue = null;

    try {
        ooxmlResult = paragraph.getOoxml();
        await context.sync();
        ooxmlValue = ooxmlResult?.value || null;
        if (ooxmlValue) return { ooxml: ooxmlValue, source: 'paragraph' };
    } catch (paragraphError) {
        console.warn(`[${logPrefix}] Paragraph.getOoxml failed, trying range.getOoxml`, paragraphError);
    }

    try {
        const paragraphRange = paragraph.getRange();
        ooxmlResult = paragraphRange.getOoxml();
        await context.sync();
        ooxmlValue = ooxmlResult?.value || null;
        if (ooxmlValue) return { ooxml: ooxmlValue, source: 'paragraphRange' };
    } catch (rangeError) {
        console.warn(`[${logPrefix}] Range.getOoxml failed for paragraph`, rangeError);
    }

    try {
        if (!propertiesPreloaded) {
            paragraph.load('parentTableCellOrNullObject, parentTableOrNullObject');
            await context.sync();
        }

        if (paragraph.parentTableCellOrNullObject && !paragraph.parentTableCellOrNullObject.isNullObject) {
            try {
                ooxmlResult = paragraph.parentTableCellOrNullObject.getOoxml();
                await context.sync();
                ooxmlValue = ooxmlResult?.value || null;
                if (ooxmlValue) return { ooxml: ooxmlValue, source: 'tableCell' };
            } catch (cellError) {
                console.warn(`[${logPrefix}] Parent table cell getOoxml failed, trying cell range`, cellError);
                try {
                    const cellRange = paragraph.parentTableCellOrNullObject.getRange();
                    ooxmlResult = cellRange.getOoxml();
                    await context.sync();
                    ooxmlValue = ooxmlResult?.value || null;
                    if (ooxmlValue) return { ooxml: ooxmlValue, source: 'tableCellRange' };
                } catch (cellRangeError) {
                    console.warn(`[${logPrefix}] Parent table cell range getOoxml failed`, cellRangeError);
                }
            }
        }

        if (!ooxmlValue && paragraph.parentTableOrNullObject && !paragraph.parentTableOrNullObject.isNullObject) {
            try {
                ooxmlResult = paragraph.parentTableOrNullObject.getOoxml();
                await context.sync();
                ooxmlValue = ooxmlResult?.value || null;
                if (ooxmlValue) return { ooxml: ooxmlValue, source: 'table' };
            } catch (tableError) {
                console.warn(`[${logPrefix}] Parent table getOoxml failed, trying table range`, tableError);
                try {
                    const tableRange = paragraph.parentTableOrNullObject.getRange();
                    ooxmlResult = tableRange.getOoxml();
                    await context.sync();
                    ooxmlValue = ooxmlResult?.value || null;
                    if (ooxmlValue) return { ooxml: ooxmlValue, source: 'tableRange' };
                } catch (tableRangeError) {
                    console.warn(`[${logPrefix}] Parent table range getOoxml failed`, tableRangeError);
                }
            }
        }
    } catch (tableTraversalError) {
        console.warn(`[${logPrefix}] Table OOXML fallback failed`, tableTraversalError);
    }

    return { ooxml: null, source: null };
}

/**
 * Inserts OOXML with `Paragraph.insertOoxml`, then retries using range-based insertion for GeneralException.
 *
 * @param {Word.Paragraph} paragraph - Target paragraph proxy
 * @param {string} wrappedOoxml - Package/fragment OOXML
 * @param {'Replace'|'After'|'Before'|'Start'|'End'} insertMode - Insert mode
 * @param {Word.RequestContext} context - Word request context
 * @param {string} [logPrefix='OOXML'] - Log prefix
 * @returns {Promise<void>}
 */
export async function insertOoxmlWithRangeFallback(paragraph, wrappedOoxml, insertMode, context, logPrefix = 'OOXML') {
    try {
        paragraph.insertOoxml(wrappedOoxml, insertMode);
        await context.sync();
        return;
    } catch (primaryError) {
        const isGeneralException = primaryError && primaryError.code === 'GeneralException';
        if (!isGeneralException) {
            throw primaryError;
        }

        console.warn(`[${logPrefix}] Paragraph.insertOoxml failed with GeneralException. Retrying via range (${insertMode}).`);

        const fallbackRange = insertMode === 'After'
            ? paragraph.getRange('End')
            : paragraph.getRange('Whole');

        fallbackRange.insertOoxml(wrappedOoxml, insertMode);
        await context.sync();
    }
}

/**
 * Runs an operation while temporarily disabling native Word track changes.
 *
 * @template T
 * @param {Word.RequestContext} context - Word request context
 * @param {(state: { originalMode: Word.ChangeTrackingMode, trackingDisabled: boolean }) => Promise<T>} operation - Operation callback
 * @param {Object} [options={}] - Toggle options
 * @param {boolean} [options.enabled=true] - Whether to disable tracking during the operation
 * @param {Word.ChangeTrackingMode|null} [options.baseTrackingMode=null] - Optional preloaded tracking mode
 * @param {string} [options.logPrefix='Tracking'] - Log prefix
 * @returns {Promise<T>}
 */
export async function withNativeTrackingDisabled(context, operation, options = {}) {
    const {
        enabled = true,
        baseTrackingMode = null,
        logPrefix = 'Tracking'
    } = options;

    let originalMode = baseTrackingMode;

    if (originalMode === null || originalMode === undefined) {
        context.document.load('changeTrackingMode');
        await context.sync();
        originalMode = context.document.changeTrackingMode;
    }

    const trackingDisabled = !!enabled && originalMode !== Word.ChangeTrackingMode.off;

    if (trackingDisabled) {
        context.document.changeTrackingMode = Word.ChangeTrackingMode.off;
        await context.sync();
    }

    try {
        return await operation({ originalMode, trackingDisabled });
    } finally {
        if (trackingDisabled) {
            context.document.changeTrackingMode = originalMode;
            await context.sync();
        } else if (enabled) {
            console.log(`[${logPrefix}] Native tracking already off; no toggle needed`);
        }
    }
}
