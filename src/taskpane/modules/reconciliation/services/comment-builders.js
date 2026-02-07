/**
 * Comment XML builders.
 */

import { NS_W, escapeXml } from '../core/types.js';

/**
 * Builds a single w:comment element.
 *
 * @param {number} commentId - Unique comment ID
 * @param {string} author - Author name
 * @param {string} content - Comment text content
 * @param {string} date - ISO date string
 * @returns {string}
 */
export function buildCommentElement(commentId, author, content, date) {
    const initials = author.split(' ').map(word => word[0]).join('').toUpperCase() || 'AI';
    const escapedContent = escapeXml(content);
    const escapedAuthor = escapeXml(author);

    return `<w:comment w:id="${commentId}" w:author="${escapedAuthor}" w:date="${date}" w:initials="${initials}">
      <w:p>
        <w:r><w:t>${escapedContent}</w:t></w:r>
      </w:p>
    </w:comment>`;
}

/**
 * Builds the complete comments.xml part.
 *
 * @param {Array<{id:number,content:string,author:string,date:string}>} comments - Placed comments
 * @returns {string}
 */
export function buildCommentsPartXml(comments) {
    if (!comments || comments.length === 0) {
        return `<w:comments xmlns:w="${NS_W}"></w:comments>`;
    }

    const commentElements = comments.map(comment =>
        buildCommentElement(comment.id, comment.author, comment.content, comment.date)
    ).join('\n    ');

    return `<w:comments xmlns:w="${NS_W}">
    ${commentElements}
  </w:comments>`;
}

/**
 * Builds inline range/reference markers for a comment id.
 *
 * @param {number} commentId - The comment ID
 * @returns {{ start: string, end: string, reference: string }}
 */
export function buildCommentMarkers(commentId) {
    return {
        start: `<w:commentRangeStart w:id="${commentId}"/>`,
        end: `<w:commentRangeEnd w:id="${commentId}"/>`,
        reference: `<w:r><w:rPr></w:rPr><w:commentReference w:id="${commentId}"/></w:r>`
    };
}
