/**
 * OOXML Reconciliation Pipeline - Patching
 *
 * Splits runs at diff boundaries and applies patch operations.
 */

import { DiffOp, RunKind } from '../core/types.js';
import { log } from '../adapters/logger.js';
import { matchListMarker, stripListMarker } from './list-markers.js';
import { serializeXml } from '../adapters/xml-adapter.js';

const XMLNS_ATTR_REGEX = /\s+xmlns:[^=]+="[^"]*"/g;

/**
 * Splits runs at diff operation boundaries for precise patching.
 *
 * @param {import('../core/types.js').RunEntry[]} runModel - Original run model
 * @param {import('../core/types.js').DiffOperation[]} diffOps - Diff operations
 * @returns {import('../core/types.js').RunEntry[]} Split run model
 */
export function splitRunsAtDiffBoundaries(runModel, diffOps) {
    const boundaries = buildSortedDiffBoundaries(diffOps);
    const newModel = [];
    let boundaryCursor = 0;

    for (const run of runModel) {
        if (run.kind !== RunKind.TEXT && run.kind !== RunKind.HYPERLINK) {
            newModel.push(run);
            continue;
        }

        while (boundaryCursor < boundaries.length && boundaries[boundaryCursor] <= run.startOffset) {
            boundaryCursor++;
        }

        let cursor = boundaryCursor;
        let currentStart = run.startOffset;
        let hasSplit = false;

        while (cursor < boundaries.length) {
            const boundary = boundaries[cursor];
            if (boundary >= run.endOffset) break;

            if (boundary > currentStart) {
                hasSplit = true;
                newModel.push({
                    ...run,
                    text: run.text.slice(currentStart - run.startOffset, boundary - run.startOffset),
                    startOffset: currentStart,
                    endOffset: boundary
                });
                currentStart = boundary;
            }
            cursor++;
        }

        boundaryCursor = cursor;

        if (!hasSplit) {
            newModel.push(run);
            continue;
        }

        newModel.push({
            ...run,
            text: run.text.slice(currentStart - run.startOffset),
            startOffset: currentStart,
            endOffset: run.endOffset
        });
    }

    return newModel;
}

/**
 * Applies diff operations to the split run model.
 *
 * @param {import('../core/types.js').RunEntry[]} splitModel - Pre-split run model
 * @param {import('../core/types.js').DiffOperation[]} diffOps - Diff operations
 * @param {Object} options - Patching options
 * @param {boolean} options.generateRedlines - Whether to generate track changes
 * @param {string} options.author - Author for track changes
 * @param {import('../core/types.js').FormatHint[]} [options.formatHints] - Format hints
 * @returns {import('../core/types.js').RunEntry[]}
 */
export function applyPatches(splitModel, diffOps, options) {
    const { generateRedlines, author } = options;
    const patchedModel = [];
    const processedInsertions = new Set();
    const diffLookup = buildPatchLookupIndex(diffOps);
    const styleLookup = buildTextRunLookup(splitModel);
    const getCoveringDiffOp = createRangeCursorLookup(diffLookup.nonInsertOps);
    const state = {
        containerStack: [],
        lastParagraphStartIndex: -1,
        currentParagraphPPrXml: '',
        currentParagraphPPrElement: null
    };

    for (const run of splitModel) {
        if (run.kind === RunKind.CONTAINER_START) {
            state.containerStack.push(run.containerId);
            patchedModel.push({ ...run });
            continue;
        }

        if (run.kind === RunKind.CONTAINER_END) {
            state.containerStack.pop();
            patchedModel.push({ ...run });
            continue;
        }

        if (run.kind === RunKind.PARAGRAPH_START) {
            state.currentParagraphPPrXml = typeof run.pPrXml === 'string' ? run.pPrXml : '';
            state.currentParagraphPPrElement = run.pPrElement || null;
            patchedModel.push({ ...run });
            state.lastParagraphStartIndex = patchedModel.length - 1;
            continue;
        }

        if (run.kind === RunKind.BOOKMARK || run.kind === RunKind.DELETION) {
            patchedModel.push({ ...run });
            continue;
        }

        const op = getCoveringDiffOp(run.startOffset, run.endOffset);
        const insertOps = diffLookup.insertOpsByStartOffset.get(run.startOffset) || [];

        for (const insertOp of insertOps) {
            if (processedInsertions.has(insertOp)) continue;
            processedInsertions.add(insertOp);
            processInsertionOperation({
                insertOp,
                splitModel,
                styleLookup,
                patchedModel,
                state,
                options,
                generateRedlines,
                author
            });
        }

        if (!op || op.type === DiffOp.EQUAL) {
            patchedModel.push({
                ...run,
                containerContext: state.containerStack.length > 0 ? state.containerStack[state.containerStack.length - 1] : null
            });
            continue;
        }

        if (op.type === DiffOp.DELETE && generateRedlines) {
            patchedModel.push({
                ...run,
                kind: RunKind.DELETION,
                author,
                containerContext: state.containerStack.length > 0 ? state.containerStack[state.containerStack.length - 1] : null
            });
        }
    }

    const endOffset = splitModel.length > 0
        ? Math.max(...splitModel.map(run => run.endOffset))
        : 0;

    for (const insertOp of diffLookup.sortedInsertOps) {
        if (insertOp.startOffset < endOffset || processedInsertions.has(insertOp)) continue;

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

function processInsertionOperation(context) {
    const {
        insertOp,
        splitModel,
        styleLookup,
        patchedModel,
        state,
        options,
        generateRedlines,
        author
    } = context;

    const lines = insertOp.text.split('\n');
    const styleSource = chooseInsertionStyle(styleLookup, insertOp.startOffset, insertOp.text);

    for (let index = 0; index < lines.length; index++) {
        const parsed = parseInsertionLine(lines[index], options.numberingService, state);
        const lineText = parsed.lineText;

        if (index > 0) {
            let newPPrXml = resolveCurrentParagraphPPrXml(state);
            if (parsed.isListLine && parsed.numId) {
                newPPrXml = options.numberingService.buildListPPr(parsed.numId, parsed.ilvl);
            }

            patchedModel.push({
                kind: RunKind.PARAGRAPH_START,
                pPrXml: newPPrXml,
                startOffset: insertOp.startOffset,
                endOffset: insertOp.startOffset,
                text: ''
            });
            state.currentParagraphPPrXml = newPPrXml;
            state.currentParagraphPPrElement = null;
            state.lastParagraphStartIndex = patchedModel.length - 1;
        } else if (parsed.isListLine && parsed.numId && state.lastParagraphStartIndex >= 0) {
            const newPPrXml = options.numberingService.buildListPPr(parsed.numId, parsed.ilvl);
            patchedModel[state.lastParagraphStartIndex].pPrXml = newPPrXml;
            patchedModel[state.lastParagraphStartIndex].pPrElement = null;
            state.currentParagraphPPrXml = newPPrXml;
            state.currentParagraphPPrElement = null;
            log(`[Patching] Converted current paragraph to list item: numId=${parsed.numId}, ilvl=${parsed.ilvl}`);
        }

        if (lineText.length > 0 || index > 0) {
            patchedModel.push({
                kind: generateRedlines ? RunKind.INSERTION : RunKind.TEXT,
                text: lineText,
                rPrXml: styleSource?.rPrXml || '',
                startOffset: insertOp.startOffset,
                endOffset: insertOp.startOffset + lineText.length,
                author: generateRedlines ? author : undefined,
                containerContext: state.containerStack.length > 0 ? state.containerStack[state.containerStack.length - 1] : null
            });
        }
    }
}

function parseInsertionLine(line, numberingService, state) {
    if (!numberingService) {
        return { lineText: line, isListLine: false, numId: null, ilvl: 0 };
    }

    const markerMatch = matchListMarker(line, { allowZeroSpaceAfterMarker: true });
    if (!markerMatch) {
        return { lineText: line, isListLine: false, numId: null, ilvl: 0 };
    }

    const marker = markerMatch[2].trim();
    const lineFormat = numberingService.detectNumberingFormat(marker);
    const indentMatch = line.match(/^(\s*)/);
    const indentSize = indentMatch ? indentMatch[1].length : 0;
    const indentStep = indentSize >= 4 ? 4 : 2;
    const indentLevel = Math.floor(indentSize / indentStep);

    const currentPPrXml = resolveCurrentParagraphPPrXml(state);
    const lineText = stripListMarker(line, { allowZeroSpaceAfterMarker: true });
    const numIdMatch = currentPPrXml.match(/w:numId w:val="(\d+)"/);
    const ilvlMatch = currentPPrXml.match(/w:ilvl w:val="(\d+)"/);
    const contextNumId = numIdMatch ? numIdMatch[1] : null;
    const contextIlvl = ilvlMatch ? parseInt(ilvlMatch[1], 10) : 0;

    const numId = numberingService.getOrCreateNumId(
        { type: lineFormat.format },
        { numId: contextNumId, type: 'unknown' }
    );

    const ilvl = lineFormat.format === 'outline'
        ? Math.min(8, lineFormat.depth)
        : Math.min(8, indentLevel + contextIlvl);

    return { lineText, isListLine: true, numId, ilvl };
}

function resolveCurrentParagraphPPrXml(state) {
    if (state.currentParagraphPPrXml) {
        return state.currentParagraphPPrXml;
    }

    if (!state.currentParagraphPPrElement) {
        return '';
    }

    state.currentParagraphPPrXml = serializeXml(state.currentParagraphPPrElement).replace(XMLNS_ATTR_REGEX, '');
    return state.currentParagraphPPrXml;
}

function chooseInsertionStyle(styleLookup, offset, insertText) {
    const prevRun = styleLookup.findRunBefore(offset);
    const nextRun = styleLookup.findRunAfter(offset);

    if (!prevRun && !nextRun) return null;
    if (!prevRun) return nextRun;
    if (!nextRun) return prevRun;

    if (insertText && insertText.endsWith(' ')) return nextRun;
    if (insertText && insertText.startsWith(' ')) return prevRun;

    return prevRun;
}

function buildTextRunLookup(runModel) {
    const textRuns = runModel.filter(run => run.kind === RunKind.TEXT);
    const starts = textRuns.map(run => run.startOffset);
    const ends = textRuns.map(run => run.endOffset);

    return {
        findRunBefore(offset) {
            let left = 0;
            let right = ends.length - 1;
            let answer = -1;

            while (left <= right) {
                const middle = (left + right) >> 1;
                if (ends[middle] <= offset) {
                    answer = middle;
                    left = middle + 1;
                } else {
                    right = middle - 1;
                }
            }

            return answer >= 0 ? textRuns[answer] : null;
        },

        findRunAfter(offset) {
            let left = 0;
            let right = starts.length - 1;
            let answer = -1;

            while (left <= right) {
                const middle = (left + right) >> 1;
                if (starts[middle] >= offset) {
                    answer = middle;
                    right = middle - 1;
                } else {
                    left = middle + 1;
                }
            }

            return answer >= 0 ? textRuns[answer] : null;
        }
    };
}

function createRangeCursorLookup(operations) {
    let cursor = 0;
    return (startOffset, endOffset) => {
        while (cursor < operations.length && operations[cursor].endOffset <= startOffset) {
            cursor++;
        }

        const operation = operations[cursor];
        if (!operation) return null;
        if (operation.startOffset <= startOffset && operation.endOffset >= endOffset) {
            return operation;
        }
        return null;
    };
}

function buildSortedDiffBoundaries(diffOps) {
    const unique = new Set();
    for (const op of diffOps) {
        unique.add(op.startOffset);
        unique.add(op.endOffset);
    }
    return Array.from(unique).sort((a, b) => a - b);
}

/**
 * Builds indexed diff lookups for patching hot paths.
 *
 * @param {import('../core/types.js').DiffOperation[]} diffOps - Diff operations
 * @returns {{
 *   insertOpsByStartOffset: Map<number, import('../core/types.js').DiffOperation[]>,
 *   nonInsertOps: import('../core/types.js').DiffOperation[],
 *   sortedInsertOps: import('../core/types.js').DiffOperation[]
 * }}
 */
function buildPatchLookupIndex(diffOps) {
    const insertOpsByStartOffset = new Map();
    const nonInsertOps = [];
    const sortedInsertOps = [];

    for (const op of diffOps) {
        if (op.type === DiffOp.INSERT) {
            if (!insertOpsByStartOffset.has(op.startOffset)) {
                insertOpsByStartOffset.set(op.startOffset, []);
            }
            insertOpsByStartOffset.get(op.startOffset).push(op);
            sortedInsertOps.push(op);
            continue;
        }
        nonInsertOps.push(op);
    }

    nonInsertOps.sort((a, b) => a.startOffset - b.startOffset || a.endOffset - b.endOffset);
    sortedInsertOps.sort((a, b) => a.startOffset - b.startOffset || a.endOffset - b.endOffset);

    return {
        insertOpsByStartOffset,
        nonInsertOps,
        sortedInsertOps
    };
}
