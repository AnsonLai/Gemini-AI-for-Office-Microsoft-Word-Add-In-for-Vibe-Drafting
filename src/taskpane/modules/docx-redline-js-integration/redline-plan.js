import { analyzeStructuredContent, getDefaultAuthor, sanitizeAiResponse } from '@ansonlai/docx-redline-js';
import { createParser } from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';
import { compileExactReplacements } from '@ansonlai/docx-redline-js/services/localized-replacement-compiler.js';
import { preprocessMarkdown } from '@ansonlai/docx-redline-js/pipeline/markdown-processor.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function createPlanningError(code, message) {
    const error = new Error(message);
    error.code = code;
    return error;
}

function normalizeAnchor(value) {
    return String(value ?? '').replace(/^\s*\[P\d+(?:\|[^\]]*)?\]\s*/, '').replace(/\s+/g, ' ').trim();
}

function paragraphAnchorMatches(change, paragraph) {
    const anchor = normalizeAnchor(change?.anchorText);
    return !anchor || normalizeAnchor(paragraph.exactText).startsWith(anchor);
}

function assertSameAppendAnchor(edit, append, paragraph) {
    if (!paragraphAnchorMatches(edit, paragraph) || !paragraphAnchorMatches(append, paragraph)) {
        throw createPlanningError('APPEND_ANCHOR_MISMATCH', 'The paragraph edit and append do not share the exact final-paragraph anchor.');
    }
}

function singleMarkdownTable(content) {
    const analysis = analyzeStructuredContent(content);
    if (!analysis.valid) {
        throw createPlanningError('INVALID_APPEND_TABLE', analysis.issues.map(issue => issue.message).join(' '));
    }
    return analysis.blocks.length === 1 && analysis.blocks[0]?.type === 'table';
}

function assertSafeTableFormatting(content, paragraph, options, fail) {
    const analysis = analyzeStructuredContent(content);
    if (!analysis.valid || !analysis.blocks.some(block => block.type === 'table')) return;
    const requestedFormatting = analysis.blocks
        .filter(block => block.type !== 'table')
        .some(block => preprocessMarkdown(block.markdown ?? block.text ?? '').formatHints.length > 0);
    if (requestedFormatting) {
        fail(
            'UNSUPPORTED_TABLE_FORMATTING',
            'The installed document engine does not reliably preserve new inline formatting in the same replacement as a Markdown table.'
        );
    }
    if (hasSameAuthorTrackedRunFormatting(options.sourceDocumentXml, paragraph.index, options.author)) {
        fail(
            'UNSUPPORTED_TABLE_FORMATTING',
            'Replacing or appending a Markdown table at a paragraph with pending same-author character-format revisions can discard that formatting in the installed document engine.'
        );
    }
}

function editedParagraphText(change, operationName, paragraph, options, decodeEscapes, contentOf) {
    const replacements = Array.isArray(change?.replacements) && change.replacements.length > 0
        ? change.replacements.map(item => ({
            find: decodeEscapes(item.find),
            replace: options.sanitizeInput === true
                ? sanitizeAiResponse(decodeEscapes(item.replace))
                : decodeEscapes(item.replace),
            ...(item.occurrence != null ? { occurrence: item.occurrence } : {})
        }))
        : operationName === 'modify_text'
            ? [{
                find: decodeEscapes(change?.originalText ?? ''),
                replace: contentOf(change, operationName)
            }]
            : null;

    if (replacements) {
        const compiled = compileExactReplacements(paragraph.exactText, replacements, 'replacements');
        if (!compiled.ok) {
            throw createPlanningError(compiled.error?.code || 'INVALID_OPERATION', compiled.error?.message || 'Localized replacement could not be composed safely.');
        }
        return compiled.desiredText;
    }
    return contentOf(change, operationName);
}

function hasSameAuthorTrackedRunFormatting(documentXml, paragraphIndex, author) {
    if (typeof documentXml !== 'string' || !documentXml || !Number.isInteger(paragraphIndex) || paragraphIndex < 1) return false;
    const parsed = createParser().parseFromString(documentXml, 'application/xml');
    const parseError = parsed.getElementsByTagName('parsererror')[0];
    if (parseError) {
        throw createPlanningError('SOURCE_XML_INSPECTION_FAILED', 'Could not inspect existing revisions before appending a table.');
    }
    const paragraph = Array.from(parsed.getElementsByTagNameNS(NS_W, 'p'))[paragraphIndex - 1];
    if (!paragraph) return false;
    const expectedAuthor = author || getDefaultAuthor();
    return Array.from(paragraph.getElementsByTagNameNS(NS_W, 'rPrChange')).some(change => (
        change.getAttribute('w:author') || change.getAttributeNS(NS_W, 'author')
    ) === expectedAuthor);
}

function coalesceFinalParagraphEditAndTableAppend(changes, paragraphs, options, contentOf, decodeEscapes) {
    if (changes.length !== 2 || paragraphs.length === 0) return null;
    const appendNames = new Set(['replace_paragraph', 'edit_paragraph']);
    const appendIndexes = changes.flatMap((change, index) => (
        Number(change?.paragraphIndex) === paragraphs.length + 1
            && appendNames.has(String(change?.operation || '').trim().toLowerCase())
            ? [index]
            : []
    ));
    if (appendIndexes.length !== 1) return null;

    const appendChangeIndex = appendIndexes[0];
    const editChangeIndex = appendChangeIndex === 0 ? 1 : 0;
    const appendChange = changes[appendChangeIndex];
    const editChange = changes[editChangeIndex];
    const appendName = String(appendChange?.operation || '').trim().toLowerCase();
    const editName = String(editChange?.operation || '').trim().toLowerCase();
    const supportedEditNames = new Set(['edit_paragraph', 'replace_paragraph', 'modify_text']);
    if (!supportedEditNames.has(editName) || Number(editChange?.paragraphIndex) !== paragraphs.length) return null;
    if (editName === 'replace_range' || editChange?.endParagraphIndex != null || appendChange?.endParagraphIndex != null) return null;
    if (Array.isArray(appendChange?.replacements) && appendChange.replacements.length > 0) return null;
    if (editName === 'edit_paragraph' && editChange?.replacements != null
        && (editChange.newContent != null || editChange.content != null)) return null;

    const paragraph = paragraphs[paragraphs.length - 1];
    if (!paragraph || paragraph.index !== paragraphs.length) return null;
    assertSameAppendAnchor(editChange, appendChange, paragraph);

    const appendedContent = contentOf(appendChange, appendName);
    if (!singleMarkdownTable(appendedContent)) return null;
    assertSafeTableFormatting(appendedContent, paragraph, options, (code, message) => {
        throw createPlanningError(code, message);
    });

    const editedText = editedParagraphText(editChange, editName, paragraph, options, decodeEscapes, contentOf);
    if (editedText.includes('\n') || editedText.includes('\r') || analyzeStructuredContent(editedText).blocks.some(block => block.type !== 'paragraph')) {
        return null;
    }
    if (preprocessMarkdown(editedText).formatHints.length > 0) {
        throw createPlanningError(
            'UNSUPPORTED_TABLE_FORMATTING',
            'The installed document engine does not preserve newly requested inline formatting when it is combined with an appended Markdown table.'
        );
    }

    return {
        operations: [{
            type: 'redline',
            target: {
                index: paragraph.index,
                exactText: String(paragraph.exactText ?? paragraph.text ?? ''),
                ...(paragraph.paragraphId ? { paragraphId: paragraph.paragraphId } : {}),
                ...(paragraph.fingerprint ? { fingerprint: paragraph.fingerprint } : {})
            },
            targetRef: `P${paragraph.index}`,
            structuredContent: true,
            modified: `${editedText}\n${appendedContent}`
        }],
        // Engine receipt operation indexes are one-based. Both model changes
        // are represented by this one canonical source-target operation.
        changeOperationIndexes: [1, 1]
    };
}

/** Build one immutable-source operation batch from the add-in's change schema. */
export function planRedlineBatchOperationsWithMapping(aiChanges, inspectedParagraphs, options = {}) {
    const paragraphs = Array.isArray(inspectedParagraphs) ? inspectedParagraphs : [];
    const changes = Array.isArray(aiChanges) ? aiChanges : [];
    const operations = [];
    const changeOperationIndexes = [];
    const decodeEscapes = value => String(value).replace(/\\n/g, '\n').replace(/\\t/g, '\t').replace(/\\r/g, '\r');
    const fail = (code, message) => {
        throw createPlanningError(code, message);
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

    const coalesced = coalesceFinalParagraphEditAndTableAppend(changes, paragraphs, options, contentOf, decodeEscapes);
    if (coalesced) return coalesced;

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
            changeOperationIndexes.push(operations.length);
            continue;
        }
        if (operationName === 'modify_text') {
            const find = decodeEscapes(change?.originalText ?? '');
            if (!find) fail('INVALID_OPERATION', 'modify_text requires originalText.');
            const replace = contentOf(change, operationName);
            operations.push({ ...common, replacements: [{ find, replace }] });
            changeOperationIndexes.push(operations.length);
            continue;
        }
        if (!['edit_paragraph', 'replace_paragraph', 'replace_range'].includes(operationName)) {
            fail('INVALID_OPERATION', `Unsupported redline operation: ${operationName || '(missing)'}.`);
        }
        const content = contentOf(change, operationName);
        assertSafeTableFormatting(content, paragraph, options, fail);
        if (append) {
            operations.push({ ...common, modified: `${target.exactText}\n${content}` });
            changeOperationIndexes.push(operations.length);
            continue;
        }
        if (operationName === 'replace_range') {
            const endIndex = Number(change?.endParagraphIndex);
            if (endIndex === requestedIndex - 1) {
                operations.push({ ...common, modified: `${content}\n${target.exactText}` });
                changeOperationIndexes.push(operations.length);
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
            changeOperationIndexes.push(operations.length);
            continue;
        }
        operations.push({ ...common, modified: content });
        changeOperationIndexes.push(operations.length);
    }
    return { operations, changeOperationIndexes };
}

/** Backward-compatible array-only planner surface. */
export function planRedlineBatchOperations(aiChanges, inspectedParagraphs, options = {}) {
    return planRedlineBatchOperationsWithMapping(aiChanges, inspectedParagraphs, options).operations;
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
