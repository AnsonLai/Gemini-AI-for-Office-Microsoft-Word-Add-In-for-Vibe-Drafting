/**
 * Shared list-marker detection/parsing helpers.
 *
 * Keeps marker parsing consistent across router, pipeline, and patching flows.
 */

const LIST_MARKER_CORE = String.raw`(?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|[-*\u2022])`;

const LINE_REGEX_STRICT = new RegExp(`^(\\s*)((?:${LIST_MARKER_CORE})\\s+)`);
const LINE_REGEX_LOOSE = new RegExp(`^(\\s*)((?:${LIST_MARKER_CORE})\\s*)`);
const MULTILINE_REGEX_STRICT = new RegExp(`^(\\s*)((?:${LIST_MARKER_CORE})\\s+)`, 'm');
const MULTILINE_REGEX_LOOSE = new RegExp(`^(\\s*)((?:${LIST_MARKER_CORE})\\s*)`, 'm');

/**
 * Determines whether text should be treated as list-target content.
 * Strict mode requires at least one whitespace after the marker.
 *
 * @param {string} text - Candidate text
 * @returns {boolean}
 */
export function isListTargetStrict(text) {
    if (typeof text !== 'string') return false;
    return text.includes('\n') && MULTILINE_REGEX_STRICT.test(text);
}

/**
 * Determines whether text should be treated as list-target content.
 * Loose mode allows markers with optional trailing whitespace.
 *
 * @param {string} text - Candidate text
 * @returns {boolean}
 */
export function isListTargetLoose(text) {
    if (typeof text !== 'string') return false;
    return text.includes('\n') && MULTILINE_REGEX_LOOSE.test(text.trim());
}

/**
 * Matches a list marker at the start of a line.
 *
 * @param {string} line - Input line
 * @param {Object} [options={}] - Match options
 * @param {boolean} [options.allowZeroSpaceAfterMarker=false] - Allow zero spaces after marker
 * @returns {RegExpMatchArray|null}
 */
export function matchListMarker(line, options = {}) {
    const { allowZeroSpaceAfterMarker = false } = options;
    const regex = allowZeroSpaceAfterMarker ? LINE_REGEX_LOOSE : LINE_REGEX_STRICT;
    return line.match(regex);
}

/**
 * Extracts the marker text from a line.
 *
 * @param {string} line - Input line
 * @param {Object} [options={}] - Match options
 * @param {boolean} [options.allowZeroSpaceAfterMarker=false] - Allow zero spaces after marker
 * @returns {string}
 */
export function extractListMarker(line, options = {}) {
    const match = matchListMarker(line, options);
    return match ? match[2].trim() : '';
}

/**
 * Strips the marker (and its immediate trailing spacing) from a line.
 *
 * @param {string} line - Input line
 * @param {Object} [options={}] - Strip options
 * @param {boolean} [options.allowZeroSpaceAfterMarker=false] - Allow zero spaces after marker
 * @returns {string}
 */
export function stripListMarker(line, options = {}) {
    const { allowZeroSpaceAfterMarker = false } = options;
    const regex = allowZeroSpaceAfterMarker ? LINE_REGEX_LOOSE : LINE_REGEX_STRICT;
    return line.replace(regex, '');
}
