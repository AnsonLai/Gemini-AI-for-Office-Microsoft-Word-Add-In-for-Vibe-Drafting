/**
 * OOXML Comment Engine
 *
 * Provides pure OOXML-based comment insertion without Word JS API calls.
 */

import { NS_W, getNextRevisionId, getRevisionTimestamp, resetRevisionIdCounter } from '../core/types.js';
import { createParser, createSerializer } from '../adapters/xml-adapter.js';
import { log, error as logError } from '../adapters/logger.js';
import { getElementsByTag, getFirstElementByTag, getXmlParseError } from '../core/xml-query.js';
import { buildCommentElement, buildCommentsPartXml, buildCommentMarkers } from './comment-builders.js';
import { createParagraphTextIndex, injectMarkersIntoParagraph } from './comment-locator.js';
import {
    injectCommentsIntoPackage as injectCommentsIntoExistingPackage,
    wrapParagraphWithComments,
    wrapWithCommentsPart
} from './comment-package.js';

export { getNextRevisionId, resetRevisionIdCounter };
export { buildCommentElement, buildCommentsPartXml, buildCommentMarkers };

/**
 * @typedef {Object} CommentRequest
 * @property {number} paragraphIndex - 1-based paragraph index
 * @property {string} textToFind - Text to attach comment to
 * @property {string} commentContent - The comment text
 */

/**
 * @typedef {Object} CommentInjectionResult
 * @property {string} oxml - Complete OOXML package with comments
 * @property {string} [commentsXml] - comments.xml content when comments are applied
 * @property {number} commentsApplied - Number of successfully placed comments
 * @property {string[]} warnings - Any issues encountered
 */

function parseDocumentOxml(oxml, parser, parseFailureWarning) {
    try {
        const xmlDoc = parser.parseFromString(oxml, 'text/xml');
        const parseError = getXmlParseError(xmlDoc);
        if (parseError) {
            return { xmlDoc: null, warning: parseFailureWarning(parseError.textContent || 'parse error') };
        }
        return { xmlDoc, warning: null };
    } catch (error) {
        return { xmlDoc: null, warning: parseFailureWarning(error.message) };
    }
}

/**
 * Injects comments into OOXML using pure XML manipulation.
 *
 * @param {string} oxml - Original document OOXML
 * @param {CommentRequest[]} comments - Comment requests
 * @param {Object} [options={}] - Options
 * @param {string} [options.author='Gemini AI'] - Author for comments
 * @returns {CommentInjectionResult}
 */
export function injectCommentsIntoOoxml(oxml, comments, options = {}) {
    const { author = 'Gemini AI' } = options;
    const date = getRevisionTimestamp();
    const warnings = [];
    const placedComments = [];

    if (!comments || comments.length === 0) {
        return {
            oxml,
            commentsApplied: 0,
            warnings: ['No comments to inject']
        };
    }

    const parser = createParser();
    const serializer = createSerializer();
    const parseResult = parseDocumentOxml(
        oxml,
        parser,
        warning => `Failed to parse OXML: ${warning}`
    );

    if (!parseResult.xmlDoc) {
        logError('[CommentEngine] Parse failure:', parseResult.warning);
        return {
            oxml,
            commentsApplied: 0,
            warnings: [parseResult.warning]
        };
    }

    const xmlDoc = parseResult.xmlDoc;
    const paragraphs = getElementsByTag(xmlDoc, 'w:p');
    log(`[CommentEngine] Found ${paragraphs.length} paragraphs, processing ${comments.length} comment requests`);

    /** @type {Array<{ paragraphIndex: number, request: CommentRequest }>} */
    const validRequests = [];
    /** @type {Map<number, number>} */
    const remainingRequestsByParagraph = new Map();

    for (const request of comments) {
        const paragraphIndex = request.paragraphIndex - 1;
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.length) {
            warnings.push(`Paragraph ${request.paragraphIndex} out of range (1-${paragraphs.length})`);
            continue;
        }

        validRequests.push({ paragraphIndex, request });
        remainingRequestsByParagraph.set(
            paragraphIndex,
            (remainingRequestsByParagraph.get(paragraphIndex) || 0) + 1
        );
    }

    /** @type {Map<number, { fullText: string, runOffsets: Array<{run: Element, start: number, end: number}> }>} */
    const paragraphIndexes = new Map();

    for (const { paragraphIndex, request } of validRequests) {
        const targetParagraph = paragraphs[paragraphIndex];
        let textIndex = paragraphIndexes.get(paragraphIndex);
        if (!textIndex) {
            textIndex = createParagraphTextIndex(targetParagraph);
            paragraphIndexes.set(paragraphIndex, textIndex);
        }

        const commentId = getNextRevisionId();
        const success = injectMarkersIntoParagraph(
            xmlDoc,
            targetParagraph,
            request.textToFind,
            commentId,
            textIndex
        );

        const remaining = (remainingRequestsByParagraph.get(paragraphIndex) || 1) - 1;
        remainingRequestsByParagraph.set(paragraphIndex, remaining);

        if (!success) {
            warnings.push(`Could not find "${request.textToFind.substring(0, 30)}..." in paragraph ${request.paragraphIndex}`);
            if (remaining === 0) {
                paragraphIndexes.delete(paragraphIndex);
            }
            continue;
        }

        placedComments.push({
            id: commentId,
            content: request.commentContent,
            author,
            date
        });
        log(`[CommentEngine] Placed comment ${commentId} on paragraph ${request.paragraphIndex}`);

        if (remaining > 0) {
            // Rebuild only when another request still targets this paragraph.
            paragraphIndexes.set(paragraphIndex, createParagraphTextIndex(targetParagraph));
        } else {
            paragraphIndexes.delete(paragraphIndex);
        }
    }

    if (placedComments.length === 0) {
        return {
            oxml,
            commentsApplied: 0,
            warnings
        };
    }

    return {
        oxml: serializer.serializeToString(xmlDoc),
        commentsXml: buildCommentsPartXml(placedComments),
        commentsApplied: placedComments.length,
        warnings
    };
}

/**
 * Injects a comment into a single paragraph OOXML and returns a complete mini-package.
 *
 * @param {string} paragraphOoxml - Paragraph OOXML (raw paragraph or pkg:package)
 * @param {string} textToFind - Target text
 * @param {string} commentContent - Comment body
 * @param {Object} [options={}] - Options
 * @param {string} [options.author='AI Assistant'] - Comment author
 * @returns {{ success: boolean, package?: string, warning?: string, commentId?: number }}
 */
export function injectCommentIntoParagraphOoxml(paragraphOoxml, textToFind, commentContent, options = {}) {
    const { author = 'AI Assistant' } = options;
    const date = getRevisionTimestamp();
    const commentId = getNextRevisionId();

    const parser = createParser();
    const serializer = createSerializer();
    const parseResult = parseDocumentOxml(
        paragraphOoxml,
        parser,
        warning => `Failed to parse paragraph OOXML: ${warning}`
    );

    if (!parseResult.xmlDoc) {
        return { success: false, warning: parseResult.warning };
    }

    const xmlDoc = parseResult.xmlDoc;
    const paragraphs = getElementsByTag(xmlDoc, 'w:p');
    if (paragraphs.length === 0) {
        return { success: false, warning: 'No paragraph found in OOXML' };
    }

    const paragraph = paragraphs[0];
    const paragraphIndex = createParagraphTextIndex(paragraph);
    const success = injectMarkersIntoParagraph(xmlDoc, paragraph, textToFind, commentId, paragraphIndex);
    if (!success) {
        return { success: false, warning: `Could not find "${textToFind.substring(0, 30)}..." in paragraph` };
    }

    const commentElement = buildCommentElement(commentId, author, commentContent, date);
    const commentsXml = `<w:comments xmlns:w="${NS_W}">${commentElement}</w:comments>`;
    const pkgPackage = getFirstElementByTag(xmlDoc, 'pkg:package');

    if (pkgPackage) {
        const withComments = injectCommentsIntoExistingPackage(serializer.serializeToString(xmlDoc), commentsXml);
        return { success: true, package: withComments, commentId };
    }

    const modifiedParagraphXml = serializer.serializeToString(xmlDoc);
    return {
        success: true,
        package: wrapParagraphWithComments(modifiedParagraphXml, commentsXml),
        commentId
    };
}

/**
 * Injects comments part into an existing OOXML package from getOoxml().
 *
 * @param {string} packageOxml - Existing pkg:package
 * @param {string} commentsXml - comments.xml payload
 * @returns {string}
 */
export function injectCommentsIntoPackage(packageOxml, commentsXml) {
    return injectCommentsIntoExistingPackage(packageOxml, commentsXml);
}

/**
 * @deprecated Use injectCommentsIntoPackage instead.
 *
 * @param {string} documentXml - Document XML
 * @param {string} commentsXml - comments.xml payload
 * @returns {string}
 */
export { wrapWithCommentsPart };
