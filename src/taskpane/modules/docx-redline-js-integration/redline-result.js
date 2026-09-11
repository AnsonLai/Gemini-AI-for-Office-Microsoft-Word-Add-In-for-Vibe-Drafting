/**
 * Normalizes the non-throwing error contract returned by docx-redline-js.
 *
 * Keep this module host-independent so callers and Node tests can share it.
 */

import { sanitizeAiResponse } from '@ansonlai/docx-redline-js';

export const INPUT_SANITIZED_WARNING = 'Input was sanitized; pass sanitizeInput: false to disable.';

export class RedlineOperationError extends Error {
    constructor(code, message, details = {}) {
        const resolvedCode = code || 'OPERATION_ERROR';
        super(`[${resolvedCode}] ${message || 'DOCX operation failed'}`);
        this.name = 'RedlineOperationError';
        this.code = resolvedCode;
        this.details = details;
    }
}

export function assertRedlineResult(result, context = 'DOCX operation') {
    if (result?.status === 'error' || result?.error) {
        const details = {
            context,
            warnings: result?.warnings || [],
            packageError: result?.error || null
        };
        for (const key of [
            'operationIndex',
            'receipt',
            'receipts',
            'rolledBack',
            'validation',
            'validationSummary'
        ]) {
            if (result?.[key] !== undefined) details[key] = result[key];
        }
        throw new RedlineOperationError(
            result?.error?.code || 'OPERATION_ERROR',
            result?.error?.message || `${context} failed`,
            details
        );
    }
    return result;
}

export function prepareOperationInput(operation, sanitizeInput = false) {
    if (
        sanitizeInput !== true
        || operation?.type !== 'redline'
        || typeof operation?.modified !== 'string'
    ) {
        return { operation, sanitized: false };
    }

    const modified = sanitizeAiResponse(operation.modified);
    if (modified === operation.modified) {
        return { operation, sanitized: false };
    }

    return {
        operation: { ...operation, modified },
        sanitized: true
    };
}
