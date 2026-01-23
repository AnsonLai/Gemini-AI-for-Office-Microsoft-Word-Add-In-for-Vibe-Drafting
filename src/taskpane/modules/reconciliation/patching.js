/**
 * OOXML Reconciliation Pipeline - Patching
 * 
 * Splits runs at diff boundaries and applies patch operations.
 */

import { DiffOp, RunKind } from './types.js';

/**
 * Splits runs at diff operation boundaries for precise patching.
 * 
 * @param {import('./types.js').RunEntry[]} runModel - Original run model
 * @param {import('./types.js').DiffOperation[]} diffOps - Diff operations
 * @returns {import('./types.js').RunEntry[]} Split run model
 */
export function splitRunsAtDiffBoundaries(runModel, diffOps) {
    // Collect all boundary points
    const boundaries = new Set();
    for (const op of diffOps) {
        boundaries.add(op.startOffset);
        boundaries.add(op.endOffset);
    }

    const newModel = [];

    for (const run of runModel) {
        // Skip non-text runs (bookmarks, containers, etc.)
        if (run.kind !== RunKind.TEXT && run.kind !== RunKind.HYPERLINK) {
            newModel.push(run);
            continue;
        }

        // Find boundaries that fall within this run
        const runBoundaries = [...boundaries]
            .filter(b => b > run.startOffset && b < run.endOffset)
            .sort((a, b) => a - b);

        if (runBoundaries.length === 0) {
            // No splits needed
            newModel.push(run);
        } else {
            // Split the run at each boundary
            let currentStart = run.startOffset;

            for (const boundary of runBoundaries) {
                const localStart = currentStart - run.startOffset;
                const localEnd = boundary - run.startOffset;

                newModel.push({
                    ...run,
                    text: run.text.slice(localStart, localEnd),
                    startOffset: currentStart,
                    endOffset: boundary
                });

                currentStart = boundary;
            }

            // Add the final segment
            const localStart = currentStart - run.startOffset;
            newModel.push({
                ...run,
                text: run.text.slice(localStart),
                startOffset: currentStart,
                endOffset: run.endOffset
            });
        }
    }

    return newModel;
}

/**
 * Style inheritance strategy - determines which run's formatting to use for insertions.
 */
const StyleInheritance = {
    /**
     * Finds the appropriate style source for an insertion.
     * 
     * @param {import('./types.js').RunEntry[]} runModel - Run model
     * @param {number} offset - Insertion offset
     * @param {string} insertText - Text being inserted
     * @returns {import('./types.js').RunEntry|null}
     */
    forInsertion(runModel, offset, insertText) {
        const prevRun = this.findRunBefore(runModel, offset);
        const nextRun = this.findRunAfter(runModel, offset);

        if (!prevRun && !nextRun) return null;
        if (!prevRun) return nextRun;
        if (!nextRun) return prevRun;

        // Text ending with space = new phrase → inherit from NEXT
        if (insertText && insertText.endsWith(' ')) return nextRun;

        // Text starting with space = continuation → inherit from PREV
        if (insertText && insertText.startsWith(' ')) return prevRun;

        // Default to previous run's style
        return prevRun;
    },

    findRunBefore(model, offset) {
        return model
            .filter(r => r.endOffset <= offset && r.kind === RunKind.TEXT)
            .pop() || null;
    },

    findRunAfter(model, offset) {
        return model.find(r => r.startOffset >= offset && r.kind === RunKind.TEXT) || null;
    }
};

/**
 * Applies diff operations to the split run model.
 * 
 * @param {import('./types.js').RunEntry[]} splitModel - Pre-split run model
 * @param {import('./types.js').DiffOperation[]} diffOps - Diff operations
 * @param {Object} options - Patching options
 * @param {boolean} options.generateRedlines - Whether to generate track changes
 * @param {string} options.author - Author for track changes
 * @param {import('./types.js').FormatHint[]} [options.formatHints] - Format hints
 * @returns {import('./types.js').RunEntry[]}
 */
export function applyPatches(splitModel, diffOps, options) {
    const { generateRedlines, author, formatHints = [] } = options;
    const patchedModel = [];
    const processedInsertions = new Set();

    const containerStack = [];

    for (const run of splitModel) {
        // Track container stack
        if (run.kind === RunKind.CONTAINER_START) {
            containerStack.push(run.containerId);
            patchedModel.push({ ...run });
            continue;
        }
        if (run.kind === RunKind.CONTAINER_END) {
            containerStack.pop();
            patchedModel.push({ ...run });
            continue;
        }

        // Bookmark pass-through
        if (run.kind === RunKind.BOOKMARK) {
            patchedModel.push({ ...run });
            continue;
        }

        // Deletions from ingestion pass through
        if (run.kind === RunKind.DELETION) {
            patchedModel.push({ ...run });
            continue;
        }

        // Find the diff operation that applies to this run
        const op = findDiffOpForRun(diffOps, run);

        // Handle insertions that occur at this run's start
        const insertOps = diffOps.filter(o =>
            o.type === DiffOp.INSERT &&
            o.startOffset === run.startOffset &&
            !processedInsertions.has(o)
        );

        for (const insertOp of insertOps) {
            processedInsertions.add(insertOp);
            const styleSource = StyleInheritance.forInsertion(splitModel, insertOp.startOffset, insertOp.text);

            patchedModel.push({
                kind: generateRedlines ? RunKind.INSERTION : RunKind.TEXT,
                text: insertOp.text,
                rPrXml: styleSource?.rPrXml || '',
                startOffset: insertOp.startOffset,
                endOffset: insertOp.startOffset + insertOp.text.length,
                author: generateRedlines ? author : undefined,
                containerContext: containerStack.length > 0 ? containerStack[containerStack.length - 1] : null
            });
        }

        // Process the run based on the diff operation
        if (!op || op.type === DiffOp.EQUAL) {
            // No change - keep the run
            patchedModel.push({
                ...run,
                containerContext: containerStack.length > 0 ? containerStack[containerStack.length - 1] : null
            });
        } else if (op.type === DiffOp.DELETE) {
            if (generateRedlines) {
                // Mark as deletion
                patchedModel.push({
                    ...run,
                    kind: RunKind.DELETION,
                    author,
                    containerContext: containerStack.length > 0 ? containerStack[containerStack.length - 1] : null
                });
            }
            // If not generating redlines, simply omit the run
        }
    }

    // Handle insertions at the very end
    const endOffset = splitModel.length > 0
        ? Math.max(...splitModel.map(r => r.endOffset))
        : 0;

    const endInsertOps = diffOps.filter(o =>
        o.type === DiffOp.INSERT &&
        o.startOffset >= endOffset &&
        !processedInsertions.has(o)
    );

    for (const insertOp of endInsertOps) {
        const lastRun = splitModel[splitModel.length - 1];

        patchedModel.push({
            kind: generateRedlines ? RunKind.INSERTION : RunKind.TEXT,
            text: insertOp.text,
            rPrXml: lastRun?.rPrXml || '',
            startOffset: insertOp.startOffset,
            endOffset: insertOp.startOffset + insertOp.text.length,
            author: generateRedlines ? author : undefined
        });
    }

    return patchedModel;
}

/**
 * Finds the diff operation that applies to a given run.
 * 
 * @param {import('./types.js').DiffOperation[]} diffOps - Diff operations
 * @param {import('./types.js').RunEntry} run - The run to check
 * @returns {import('./types.js').DiffOperation|null}
 */
function findDiffOpForRun(diffOps, run) {
    // Find an operation that covers this run
    for (const op of diffOps) {
        if (op.type === DiffOp.INSERT) continue; // Insertions handled separately

        if (op.startOffset <= run.startOffset && op.endOffset >= run.endOffset) {
            return op;
        }
    }
    return null;
}
