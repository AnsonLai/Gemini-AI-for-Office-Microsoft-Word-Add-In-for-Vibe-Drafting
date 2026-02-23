import { applyWordOperation } from './word-operation-runner.js';
import { toScopedSharedRedlineOperation } from '@ansonlai/docx-redline-js/orchestration/redline-operation-converter.js';

function normalizeNeedleText(value) {
    if (value == null) return '';
    return String(value)
        .replace(/\\n/g, '\n')
        .replace(/\\t/g, '\t')
        .replace(/\\r/g, '\r')
        .trim();
}

export function findNearbyParagraphIndexForModifyText(paragraphItems, startIndex, change) {
    const items = Array.isArray(paragraphItems) ? paragraphItems : [];
    if (!Number.isInteger(startIndex) || startIndex < 0 || startIndex >= items.length) {
        return startIndex;
    }

    const originalText = normalizeNeedleText(change?.originalText);
    if (!originalText) {
        return startIndex;
    }

    const textAt = index => String(items[index]?.text ?? '');
    const startText = textAt(startIndex);
    if (startText.trim().length > 0) {
        return startIndex;
    }

    const originalLower = originalText.toLowerCase();
    const maxDistance = 12;
    let bestExact = null;
    let bestInsensitive = null;

    for (let distance = 1; distance <= maxDistance; distance += 1) {
        const candidates = [startIndex - distance, startIndex + distance];
        for (const candidateIndex of candidates) {
            if (candidateIndex < 0 || candidateIndex >= items.length) continue;

            const candidateText = textAt(candidateIndex);
            if (!candidateText.trim()) continue;

            if (candidateText.includes(originalText)) {
                if (bestExact == null) {
                    bestExact = candidateIndex;
                }
                continue;
            }

            if (candidateText.toLowerCase().includes(originalLower)) {
                if (bestInsensitive == null) {
                    bestInsensitive = candidateIndex;
                }
            }
        }

        if (bestExact != null) return bestExact;
    }

    if (bestInsensitive != null) return bestInsensitive;
    return startIndex;
}

function parseParagraphIndex(value) {
    const parsed = Number.parseInt(String(value ?? ''), 10);
    return Number.isInteger(parsed) ? parsed - 1 : null;
}

export async function applyRedlineChangesToWordContext(context, aiChanges, options = {}) {
    const changes = Array.isArray(aiChanges) ? aiChanges : [];
    const logPrefix = options.logPrefix || 'Redline/Shared';
    const onInfo = typeof options.onInfo === 'function'
        ? options.onInfo
        : message => console.log(`[${logPrefix}] ${message}`);
    const onWarn = typeof options.onWarn === 'function'
        ? options.onWarn
        : message => console.warn(`[${logPrefix}] ${message}`);

    let changesApplied = 0;

    for (const change of changes) {
        try {
            const operationName = String(change?.operation || '').trim().toLowerCase();
            const startIndex = parseParagraphIndex(change?.paragraphIndex);
            if (!Number.isInteger(startIndex) || startIndex < 0) {
                onWarn(`Invalid start paragraph index: ${change?.paragraphIndex}`);
                continue;
            }

            const paragraphs = context.document.body.paragraphs;
            paragraphs.load('items/text');
            await context.sync();

            const paragraphCount = paragraphs.items.length;
            if (startIndex >= paragraphCount) {
                onWarn(`Out-of-range target P${change?.paragraphIndex} (count=${paragraphCount}); no-op.`);
                continue;
            }

            let effectiveStartIndex = startIndex;
            if (operationName === 'modify_text') {
                effectiveStartIndex = findNearbyParagraphIndexForModifyText(paragraphs.items, startIndex, change);
                if (effectiveStartIndex !== startIndex) {
                    onWarn(
                        `Rebased modify_text from P${startIndex + 1} to P${effectiveStartIndex + 1} based on originalText match.`
                    );
                }
            }

            let endIndex = effectiveStartIndex;
            let insertionBeforeStart = false;
            if (operationName === 'replace_range') {
                const requestedEndIndex = parseParagraphIndex(change?.endParagraphIndex);
                if (!Number.isInteger(requestedEndIndex) || requestedEndIndex < -1) {
                    onWarn(`Invalid replace_range endParagraphIndex: ${change?.endParagraphIndex}; no-op.`);
                    continue;
                }

                if (requestedEndIndex === effectiveStartIndex - 1) {
                    insertionBeforeStart = true;
                    endIndex = effectiveStartIndex;
                    onWarn(
                        `Normalizing replace_range insertion-before-target (P${effectiveStartIndex + 1}..P${requestedEndIndex + 1}) to scoped insertion.`
                    );
                } else {
                    endIndex = requestedEndIndex;
                    if (endIndex < effectiveStartIndex || endIndex >= paragraphCount) {
                        onWarn(`Invalid replace_range endParagraphIndex: ${change?.endParagraphIndex}; no-op.`);
                        continue;
                    }
                }
            }

            const startParagraph = paragraphs.items[effectiveStartIndex];
            const scopeParagraphCount = insertionBeforeStart
                ? 1
                : (endIndex - effectiveStartIndex) + 1;

            const converted = toScopedSharedRedlineOperation(change, {
                scopeStartText: startParagraph.text || '',
                scopeParagraphCount,
                insertionBeforeStart
            });
            if (!converted.ok) {
                onWarn(`Skipping change: ${converted.reason}`);
                continue;
            }

            const applied = await applyWordOperation(
                context,
                converted.operation,
                scopeParagraphCount === 1
                    ? { paragraph: startParagraph }
                    : { paragraph: startParagraph, endParagraph: paragraphs.items[endIndex] },
                {
                    author: options.author,
                    generateRedlines: options.generateRedlines,
                    disableNativeTracking: options.disableNativeTracking,
                    baseTrackingMode: options.baseTrackingMode ?? null,
                    logPrefix,
                    onInfo,
                    onWarn
                }
            );

            if (applied) {
                changesApplied += 1;
            } else {
                onWarn(`No changes produced for change: ${JSON.stringify(change)}`);
            }
        } catch (changeError) {
            onWarn(`Failed to apply change ${JSON.stringify(change)}: ${changeError?.message || changeError}`);
        }
    }

    onInfo(`Total changes applied: ${changesApplied}`);
    return { changesApplied };
}
