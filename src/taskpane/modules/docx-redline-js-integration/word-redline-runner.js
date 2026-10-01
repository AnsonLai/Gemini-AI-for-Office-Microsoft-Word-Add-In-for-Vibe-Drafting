import { executePureOoxmlBatch } from './word-operation-runner.js';
import { sanitizeAiResponse } from '@ansonlai/docx-redline-js';

/** Build one immutable-source operation batch from the add-in's change schema. */
export function planRedlineBatchOperations(aiChanges, inspectedParagraphs, options = {}) {
    const paragraphs = Array.isArray(inspectedParagraphs) ? inspectedParagraphs : [];
    const changes = Array.isArray(aiChanges) ? aiChanges : [];
    const operations = [];
    const decodeEscapes = value => String(value).replace(/\\n/g, '\n').replace(/\\t/g, '\t').replace(/\\r/g, '\r');
    const fail = (code, message) => {
        const error = new Error(message);
        error.code = code;
        throw error;
    };
    const paragraphAt = number => {
        const index = Number(number);
        if (!Number.isInteger(index) || index < 1 || index > paragraphs.length) {
            fail('TARGET_NOT_FOUND', `Paragraph P${number} is outside the inspected document.`);
        }
        return paragraphs[index - 1];
    };
    const descriptor = paragraph => ({
        index: paragraph.index,
        exactText: String(paragraph.exactText ?? paragraph.text ?? ''),
        ...(paragraph.paragraphId ? { paragraphId: paragraph.paragraphId } : {}),
        ...(paragraph.fingerprint ? { fingerprint: paragraph.fingerprint } : {})
    });
    const contentOf = (change, operationName) => {
        const value = operationName === 'edit_paragraph'
            ? (change.newContent ?? change.content ?? change.replacementText)
            : operationName === 'modify_text'
                ? (change.replacementText ?? change.content ?? change.newContent)
                : (change.content ?? change.newContent ?? change.replacementText);
        if (value == null) fail('INVALID_OPERATION', `${operationName} requires replacement content.`);
        const decoded = decodeEscapes(value);
        return options.sanitizeInput === true ? sanitizeAiResponse(decoded) : decoded;
    };
    for (const change of changes) {
        const operationName = String(change?.operation || '').trim().toLowerCase();
        const requestedIndex = Number(change?.paragraphIndex);
        const append = requestedIndex === paragraphs.length + 1
            && ['replace_paragraph', 'replace_range', 'edit_paragraph'].includes(operationName);
        const paragraph = paragraphAt(append ? paragraphs.length : requestedIndex);
        const target = descriptor(paragraph);
        const common = { type: 'redline', target, targetRef: `P${paragraph.index}`, structuredContent: true };
        if (Array.isArray(change?.replacements) && change.replacements.length > 0) {
            if (operationName === 'replace_range') fail('INVALID_OPERATION', 'Localized replacements require one paragraph.');
            operations.push({
                ...common,
                replacements: change.replacements.map(item => ({
                    find: decodeEscapes(item.find),
                    replace: options.sanitizeInput === true
                        ? sanitizeAiResponse(decodeEscapes(item.replace))
                        : decodeEscapes(item.replace),
                    ...(item.occurrence ? { occurrence: item.occurrence } : {})
                }))
            });
            continue;
        }
        if (operationName === 'modify_text') {
            const find = decodeEscapes(change?.originalText ?? '');
            if (!find) fail('INVALID_OPERATION', 'modify_text requires originalText.');
            const replace = contentOf(change, operationName);
            operations.push({ ...common, replacements: [{ find, replace }] });
            continue;
        }
        if (!['edit_paragraph', 'replace_paragraph', 'replace_range'].includes(operationName)) {
            fail('INVALID_OPERATION', `Unsupported redline operation: ${operationName || '(missing)'}.`);
        }
        const content = contentOf(change, operationName);
        if (append) {
            operations.push({ ...common, modified: `${target.exactText}\n${content}` });
            continue;
        }
        if (operationName === 'replace_range') {
            const endIndex = Number(change?.endParagraphIndex);
            if (endIndex === requestedIndex - 1) {
                operations.push({ ...common, modified: `${content}\n${target.exactText}` });
                continue;
            }
            if (!Number.isInteger(endIndex) || endIndex < requestedIndex || endIndex > paragraphs.length) {
                fail('TARGET_NOT_FOUND', `Invalid replace_range end P${change?.endParagraphIndex}.`);
            }
            operations.push({
                ...common,
                targetEndRef: `P${endIndex}`,
                modified: content
            });
            continue;
        }
        operations.push({ ...common, modified: content });
    }
    return operations;
}

/** Apply every proposed edit against one immutable OOXML snapshot and write once. */
export async function applyRedlineChangesToWordContext(context, aiChanges, options = {}) {
    const changes = Array.isArray(aiChanges) ? aiChanges : [];
    if (changes.length === 0) return { changesApplied: 0, skipped: [], written: false, writeAttempted: false, mutationOutcome: 'noop' };
    const logPrefix = options.logPrefix || 'Redline/Shared';
    const onInfo = options.onInfo || (() => {});
    const onWarn = options.onWarn || (() => console.warn(`[${logPrefix}] Batch warning; consult the structured result.`));
    const batchRunner = options.batchRunner || executePureOoxmlBatch;
    try {
        const result = await batchRunner(
            context,
            context.document.body,
            source => planRedlineBatchOperations(changes, source.paragraphs, options),
            {
                author: options.author,
                generateRedlines: options.generateRedlines,
                sanitizeInput: options.sanitizeInput === true,
                disableNativeTracking: options.disableNativeTracking,
                baseTrackingMode: options.baseTrackingMode ?? null,
                onInfo,
                onWarn
            }
        );
        const receipts = Array.isArray(result?.receipts) ? result.receipts : [];
        const changesApplied = result?.written === true
            ? receipts.filter(receipt => receipt.committed === true).length
            : 0;
        const skipped = changes.flatMap((change, index) => {
            const receipt = receipts.find(item => item.operationIndex === index + 1);
            if (result?.written === true && receipt?.committed) return [];
            const item = result?.results?.[index];
            const error = item?.error || result?.error;
            return [{
                paragraphIndex: change?.paragraphIndex,
                operation: change?.operation,
                reason: error?.message || (result?.hasChanges === false ? 'no changes produced' : 'batch was rolled back'),
                ...(error?.code ? { code: error.code } : {}),
                ...(receipt ? { receipt } : {}),
                ...(result?.rolledBack !== undefined ? { rolledBack: result.rolledBack } : {})
            }];
        });
        if (result?.status === 'error') onWarn(result.error?.message || 'Redline batch failed.');
        onInfo(`Total changes applied: ${changesApplied}`);
        return {
            changesApplied, skipped, batchResult: result,
            status: result?.status, error: result?.error,
            receipts, written: result?.written === true,
            writeAttempted: result?.writeAttempted === true,
            mutationOutcome: result?.mutationOutcome || (result?.written ? 'applied' : result?.rolledBack ? 'rolled_back' : result?.hasChanges === false ? 'noop' : 'refused')
        };
    } catch (error) {
        onWarn(`Redline batch failed: ${error?.message || error}`);
        return {
            changesApplied: 0,
            skipped: changes.map(change => ({
                paragraphIndex: change?.paragraphIndex,
                operation: change?.operation,
                reason: error?.message || String(error),
                ...(error?.code ? { code: error.code } : {})
            })),
            error,
            written: false, writeAttempted: false, mutationOutcome: 'failed'
        };
    }
}
