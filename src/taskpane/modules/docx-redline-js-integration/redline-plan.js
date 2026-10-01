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

function staleDocumentContext(message) {
    const error = new Error(message);
    error.code = 'STALE_DOCUMENT_CONTEXT';
    return error;
}

/** Refuse edits whose canonical source paragraphs changed after context creation. */
export function assertSourceBaseline(changes, liveParagraphs, sourceBaseline) {
    if (!Array.isArray(sourceBaseline) || sourceBaseline.length === 0) {
        throw staleDocumentContext('The source baseline is unavailable; reread the document before editing.');
    }

    const live = Array.isArray(liveParagraphs) ? liveParagraphs : [];
    const baseline = new Map();
    for (const paragraph of sourceBaseline) {
        if (!paragraph || !Number.isInteger(paragraph.index) || paragraph.index < 1
            || typeof paragraph.exactText !== 'string' || typeof paragraph.fingerprint !== 'string'
            || paragraph.fingerprint.length === 0 || baseline.has(paragraph.index)) {
            throw staleDocumentContext('The source baseline is incomplete or malformed; reread the document before editing.');
        }
        baseline.set(paragraph.index, paragraph);
    }

    const compareParagraph = index => {
        const expected = baseline.get(index);
        const actual = live[index - 1];
        if (!expected || !actual || actual.index !== index
            || typeof actual.exactText !== 'string'
            || typeof actual.fingerprint !== 'string'
            || actual.exactText !== expected.exactText
            || actual.fingerprint !== expected.fingerprint) {
            throw staleDocumentContext(`Paragraph P${index} changed since the edit context was created; reread before applying this batch.`);
        }
    };

    for (const change of changes) {
        const operation = String(change?.operation || '').trim().toLowerCase();
        const startIndex = change?.paragraphIndex;
        if (!Number.isInteger(startIndex) || startIndex < 1) {
            throw staleDocumentContext('The edit target could not be matched to its source baseline.');
        }

        const append = startIndex === sourceBaseline.length + 1
            && ['replace_paragraph', 'replace_range', 'edit_paragraph'].includes(operation);
        if (append) {
            if (live.length !== sourceBaseline.length || sourceBaseline.length === 0) {
                throw staleDocumentContext('The document paragraph count changed since the append context was created; reread before applying this batch.');
            }
            compareParagraph(sourceBaseline.length);
            continue;
        }

        if (operation === 'replace_range') {
            const endIndex = change?.endParagraphIndex;
            if (endIndex === startIndex - 1) {
                // This schema form inserts before the original start paragraph.
                compareParagraph(startIndex);
                continue;
            }
            if (!Number.isInteger(endIndex) || endIndex < startIndex || endIndex > sourceBaseline.length) {
                throw staleDocumentContext('The replacement range could not be matched to its source baseline.');
            }
            for (let index = startIndex; index <= endIndex; index += 1) compareParagraph(index);
            continue;
        }

        compareParagraph(startIndex);
    }
}

/** Apply every proposed edit against one immutable OOXML snapshot and write once. */
