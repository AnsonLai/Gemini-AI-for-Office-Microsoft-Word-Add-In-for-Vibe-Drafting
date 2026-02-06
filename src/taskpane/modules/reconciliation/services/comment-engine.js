/**
 * OOXML Comment Engine
 * 
 * Provides pure OOXML-based comment insertion, avoiding the Word JS API.
 * 
 * OOXML Comments Structure:
 * 1. comments.xml part - Contains all w:comment elements
 * 2. Inline markers in document.xml - commentRangeStart, commentRangeEnd, commentReference
 * 3. Relationship entry linking to comments.xml
 */

import { NS_W, escapeXml, getNextRevisionId, resetRevisionIdCounter } from '../core/types.js';
import { createParser, createSerializer } from '../adapters/xml-adapter.js';
import { log, error as logError } from '../adapters/logger.js';

export { getNextRevisionId, resetRevisionIdCounter };

// ============================================================================
// TYPES
// ============================================================================

/**
 * @typedef {Object} CommentRequest
 * @property {number} paragraphIndex - 1-based paragraph index
 * @property {string} textToFind - Text to attach comment to
 * @property {string} commentContent - The comment text
 */

/**
 * @typedef {Object} CommentInjectionResult
 * @property {string} oxml - Complete OOXML package with comments
 * @property {number} commentsApplied - Number of successfully placed comments
 * @property {string[]} warnings - Any issues encountered
 */

/**
 * @typedef {Object} PlacedComment
 * @property {number} id - Unique comment ID
 * @property {string} content - Comment content
 * @property {string} author - Comment author
 * @property {string} date - ISO date string
 */

// ============================================================================
// OOXML BUILDERS
// ============================================================================

/**
 * Builds a single w:comment element.
 * 
 * @param {number} commentId - Unique comment ID
 * @param {string} author - Author name
 * @param {string} content - Comment text content
 * @param {string} date - ISO date string
 * @returns {string} w:comment XML string
 */
export function buildCommentElement(commentId, author, content, date) {
    const initials = author.split(' ').map(w => w[0]).join('').toUpperCase() || 'AI';
    const escapedContent = escapeXml(content);
    const escapedAuthor = escapeXml(author);

    return `<w:comment w:id="${commentId}" w:author="${escapedAuthor}" w:date="${date}" w:initials="${initials}">
      <w:p>
        <w:r><w:t>${escapedContent}</w:t></w:r>
      </w:p>
    </w:comment>`;
}

/**
 * Builds the complete comments.xml part content.
 * 
 * @param {PlacedComment[]} comments - Array of placed comments
 * @returns {string} Complete w:comments XML
 */
export function buildCommentsPartXml(comments) {
    if (!comments || comments.length === 0) {
        return `<w:comments xmlns:w="${NS_W}"></w:comments>`;
    }

    const commentElements = comments.map(c =>
        buildCommentElement(c.id, c.author, c.content, c.date)
    ).join('\n    ');

    return `<w:comments xmlns:w="${NS_W}">
    ${commentElements}
  </w:comments>`;
}

/**
 * Builds the inline markers for a comment.
 * These are inserted around the commented text range.
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

// ============================================================================
// TEXT LOCATION AND MARKER INJECTION
// ============================================================================

/**
 * Finds text within paragraph runs and returns character offsets.
 * 
 * @param {Element} paragraph - w:p element
 * @param {string} searchText - Text to find
 * @returns {{ found: boolean, startRun?: Element, startOffset?: number, endRun?: Element, endOffset?: number, runs?: Element[] }}
 */
function findTextInParagraph(paragraph, searchText) {
    const runs = Array.from(paragraph.getElementsByTagName('w:r'));
    let fullText = '';
    const runOffsets = [];

    // Build full text and track run boundaries
    for (const run of runs) {
        const start = fullText.length;
        const textNodes = run.getElementsByTagName('w:t');
        for (const t of Array.from(textNodes)) {
            fullText += t.textContent || '';
        }
        runOffsets.push({ run, start, end: fullText.length });
    }

    // Search for text
    const searchIndex = fullText.indexOf(searchText);
    if (searchIndex === -1) {
        return { found: false };
    }

    const searchEnd = searchIndex + searchText.length;

    // Find which runs contain the start and end
    let startRun = null, endRun = null;
    let startOffset = 0, endOffset = 0;
    const affectedRuns = [];

    for (const { run, start, end } of runOffsets) {
        if (searchIndex >= start && searchIndex < end) {
            startRun = run;
            startOffset = searchIndex - start;
        }
        if (searchEnd > start && searchEnd <= end) {
            endRun = run;
            endOffset = searchEnd - start;
        }
        if (searchIndex < end && searchEnd > start) {
            affectedRuns.push(run);
        }
    }

    return {
        found: true,
        startRun,
        startOffset,
        endRun,
        endOffset,
        runs: affectedRuns
    };
}

/**
 * Injects comment markers into a paragraph around specified text.
 * Modifies the DOM in-place.
 * 
 * This version properly splits runs at text boundaries to ensure
 * only the exact target text is highlighted, not entire runs.
 * 
 * @param {Document} xmlDoc - The XML document
 * @param {Element} paragraph - The w:p element
 * @param {string} textToFind - Text to attach comment to
 * @param {number} commentId - The comment ID
 * @returns {boolean} True if markers were successfully injected
 */
function injectMarkersIntoParagraph(xmlDoc, paragraph, textToFind, commentId) {
    const location = findTextInParagraph(paragraph, textToFind);
    if (!location.found || !location.startRun) {
        return false;
    }

    // Create marker elements
    const startMarker = xmlDoc.createElementNS(NS_W, 'w:commentRangeStart');
    startMarker.setAttribute('w:id', String(commentId));

    const endMarker = xmlDoc.createElementNS(NS_W, 'w:commentRangeEnd');
    endMarker.setAttribute('w:id', String(commentId));

    // Create comment reference run
    const refRun = xmlDoc.createElementNS(NS_W, 'w:r');
    const refEl = xmlDoc.createElementNS(NS_W, 'w:commentReference');
    refEl.setAttribute('w:id', String(commentId));
    refRun.appendChild(refEl);

    // Handle the case where start and end are in the same run
    if (location.startRun === location.endRun) {
        // Split the run into: [before][highlighted][after]
        const run = location.startRun;
        const textNode = run.getElementsByTagName('w:t')[0];
        if (!textNode) {
            // Fallback: place markers around the run
            run.parentNode.insertBefore(startMarker, run);
            if (run.nextSibling) {
                run.parentNode.insertBefore(endMarker, run.nextSibling);
                endMarker.parentNode.insertBefore(refRun, endMarker.nextSibling);
            } else {
                run.parentNode.appendChild(endMarker);
                run.parentNode.appendChild(refRun);
            }
            return true;
        }

        const fullText = textNode.textContent || '';
        const beforeText = fullText.substring(0, location.startOffset);
        const highlightText = fullText.substring(location.startOffset, location.endOffset);
        const afterText = fullText.substring(location.endOffset);

        // Clone the run properties (w:rPr) if it exists
        const rPr = run.getElementsByTagName('w:rPr')[0];

        // Build replacement content
        const parent = run.parentNode;

        // Create "before" run if there's text before
        if (beforeText) {
            const beforeRun = cloneRunWithText(xmlDoc, run, rPr, beforeText);
            parent.insertBefore(beforeRun, run);
        }

        // Insert start marker
        parent.insertBefore(startMarker, run);

        // Replace the original run with just the highlighted text
        textNode.textContent = highlightText;

        // Insert end marker and reference after this run
        if (run.nextSibling) {
            parent.insertBefore(endMarker, run.nextSibling);
        } else {
            parent.appendChild(endMarker);
        }
        parent.insertBefore(refRun, endMarker.nextSibling || null);

        // Create "after" run if there's text after
        if (afterText) {
            const afterRun = cloneRunWithText(xmlDoc, run, rPr, afterText);
            parent.insertBefore(afterRun, refRun.nextSibling || null);
        }
    } else {
        // Text spans multiple runs - handle start and end runs separately

        // Handle start run: split if needed and place start marker
        const startTextNode = location.startRun.getElementsByTagName('w:t')[0];
        if (startTextNode && location.startOffset > 0) {
            const fullText = startTextNode.textContent || '';
            const beforeText = fullText.substring(0, location.startOffset);
            const highlightText = fullText.substring(location.startOffset);

            if (beforeText) {
                const rPr = location.startRun.getElementsByTagName('w:rPr')[0];
                const beforeRun = cloneRunWithText(xmlDoc, location.startRun, rPr, beforeText);
                location.startRun.parentNode.insertBefore(beforeRun, location.startRun);
            }
            startTextNode.textContent = highlightText;
        }

        // Place start marker before the start run
        location.startRun.parentNode.insertBefore(startMarker, location.startRun);

        // Handle end run: split if needed and place end marker
        const endRun = location.endRun || location.startRun;
        const endTextNode = endRun.getElementsByTagName('w:t')[0];
        if (endTextNode && location.endOffset < (endTextNode.textContent || '').length) {
            const fullText = endTextNode.textContent || '';
            const highlightText = fullText.substring(0, location.endOffset);
            const afterText = fullText.substring(location.endOffset);

            endTextNode.textContent = highlightText;

            if (afterText) {
                const rPr = endRun.getElementsByTagName('w:rPr')[0];
                const afterRun = cloneRunWithText(xmlDoc, endRun, rPr, afterText);
                if (endRun.nextSibling) {
                    endRun.parentNode.insertBefore(afterRun, endRun.nextSibling);
                } else {
                    endRun.parentNode.appendChild(afterRun);
                }
            }
        }

        // Place end marker and reference after the end run
        if (endRun.nextSibling) {
            endRun.parentNode.insertBefore(endMarker, endRun.nextSibling);
            endMarker.parentNode.insertBefore(refRun, endMarker.nextSibling);
        } else {
            endRun.parentNode.appendChild(endMarker);
            endRun.parentNode.appendChild(refRun);
        }
    }

    return true;
}

/**
 * Creates a clone of a run with new text content.
 * Preserves the run properties (formatting).
 */
function cloneRunWithText(xmlDoc, originalRun, rPr, newText) {
    const newRun = xmlDoc.createElementNS(NS_W, 'w:r');

    // Clone run properties if they exist
    if (rPr) {
        newRun.appendChild(rPr.cloneNode(true));
    }

    // Create new text element
    const newT = xmlDoc.createElementNS(NS_W, 'w:t');
    // Preserve spaces
    newT.setAttribute('xml:space', 'preserve');
    newT.textContent = newText;
    newRun.appendChild(newT);

    return newRun;
}

// ============================================================================
// MAIN EXPORT
// ============================================================================

/**
 * Injects comments into OOXML using pure XML manipulation.
 * 
 * @param {string} oxml - Original document OOXML
 * @param {CommentRequest[]} comments - Array of comment requests
 * @param {Object} [options={}] - Options
 * @param {string} [options.author='Gemini AI'] - Author for comments
 * @returns {CommentInjectionResult}
 */
export function injectCommentsIntoOoxml(oxml, comments, options = {}) {
    const { author = 'Gemini AI' } = options;
    const date = new Date().toISOString();
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

    let xmlDoc;
    try {
        xmlDoc = parser.parseFromString(oxml, 'text/xml');
    } catch (e) {
        logError('[CommentEngine] Failed to parse OXML:', e);
        return {
            oxml,
            commentsApplied: 0,
            warnings: [`Failed to parse OXML: ${e.message}`]
        };
    }

    // Check for parse errors
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        logError('[CommentEngine] XML parse error:', parseError.textContent);
        return {
            oxml,
            commentsApplied: 0,
            warnings: ['XML parse error in document']
        };
    }

    // Get all paragraphs
    const paragraphs = Array.from(xmlDoc.getElementsByTagName('w:p'));
    log(`[CommentEngine] Found ${paragraphs.length} paragraphs, processing ${comments.length} comment requests`);

    // Process each comment request
    for (const comment of comments) {
        const pIndex = comment.paragraphIndex - 1; // Convert to 0-based

        if (pIndex < 0 || pIndex >= paragraphs.length) {
            warnings.push(`Paragraph ${comment.paragraphIndex} out of range (1-${paragraphs.length})`);
            continue;
        }

        const targetParagraph = paragraphs[pIndex];
        const commentId = getNextRevisionId();

        const success = injectMarkersIntoParagraph(
            xmlDoc,
            targetParagraph,
            comment.textToFind,
            commentId
        );

        if (success) {
            placedComments.push({
                id: commentId,
                content: comment.commentContent,
                author,
                date
            });
            log(`[CommentEngine] Placed comment ${commentId} on paragraph ${comment.paragraphIndex}`);
        } else {
            warnings.push(`Could not find "${comment.textToFind.substring(0, 30)}..." in paragraph ${comment.paragraphIndex}`);
        }
    }

    if (placedComments.length === 0) {
        return {
            oxml,
            commentsApplied: 0,
            warnings
        };
    }

    // Serialize the modified document
    const modifiedOxml = serializer.serializeToString(xmlDoc);

    // Build the comments.xml content
    const commentsXml = buildCommentsPartXml(placedComments);

    return {
        oxml: modifiedOxml,
        commentsXml,
        commentsApplied: placedComments.length,
        warnings
    };
}

/**
 * Injects a comment into a single paragraph's OOXML and returns a complete
 * mini-package suitable for paragraph.insertOoxml().
 * 
 * This is the surgical approach - instead of replacing the entire document,
 * we just replace the affected paragraph. Since the text content is unchanged
 * (only comment markers are added), Word should not generate redlines.
 * 
 * @param {string} paragraphOoxml - The paragraph's OOXML from paragraph.getOoxml()
 * @param {string} textToFind - Text to attach the comment to
 * @param {string} commentContent - The comment text
 * @param {Object} [options={}] - Options
 * @param {string} [options.author='AI Assistant'] - Author for the comment
 * @returns {{ success: boolean, package?: string, warning?: string }}
 */
export function injectCommentIntoParagraphOoxml(paragraphOoxml, textToFind, commentContent, options = {}) {
    const { author = 'AI Assistant' } = options;
    const date = new Date().toISOString();
    const commentId = getNextRevisionId();

    const parser = createParser();
    const serializer = createSerializer();

    let xmlDoc;
    try {
        xmlDoc = parser.parseFromString(paragraphOoxml, 'text/xml');
    } catch (e) {
        return { success: false, warning: `Failed to parse paragraph OOXML: ${e.message}` };
    }

    // Check for parse errors
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        return { success: false, warning: 'XML parse error in paragraph' };
    }

    // Find the w:p element (paragraph) - it might be inside a pkg:package or directly
    const paragraphs = xmlDoc.getElementsByTagName('w:p');
    if (paragraphs.length === 0) {
        return { success: false, warning: 'No paragraph found in OOXML' };
    }

    // Use the first paragraph (there should only be one for paragraph.getOoxml())
    const paragraph = paragraphs[0];

    // Inject comment markers
    const success = injectMarkersIntoParagraph(xmlDoc, paragraph, textToFind, commentId);

    if (!success) {
        return { success: false, warning: `Could not find "${textToFind.substring(0, 30)}..." in paragraph` };
    }

    // Build the comment entry
    const commentElement = buildCommentElement(commentId, author, commentContent, date);
    const commentsXml = `<w:comments xmlns:w="${NS_W}">${commentElement}</w:comments>`;

    // Check if this is already a pkg:package (from paragraph.getOoxml()) or raw XML
    const pkgPackage = xmlDoc.getElementsByTagName('pkg:package')[0];

    if (pkgPackage) {
        // It's already a package - inject the comments part into it
        const result = injectCommentsIntoPackage(serializer.serializeToString(xmlDoc), commentsXml);
        return { success: true, package: result, commentId };
    } else {
        // Raw XML - wrap it in a minimal package with comments
        const modifiedParagraphXml = serializer.serializeToString(xmlDoc);
        const wrappedPackage = wrapParagraphWithComments(modifiedParagraphXml, commentsXml);
        return { success: true, package: wrappedPackage, commentId };
    }
}

/**
 * Wraps a paragraph's XML with a minimal package structure including comments.
 * 
 * @param {string} paragraphXml - The modified paragraph XML with comment markers
 * @param {string} commentsXml - The comments.xml content
 * @returns {string} Complete pkg:package for insertOoxml
 */
function wrapParagraphWithComments(paragraphXml, commentsXml) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml">
    <pkg:xmlData>
      ${commentsXml}
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          ${paragraphXml}
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

/**
 * Injects comments part into an existing OOXML package from getOoxml().
 * 
 * The package from getOoxml() is a pkg:package with pkg:part elements.
 * This function:
 * 1. Adds the comments.xml part to the package
 * 2. Updates the document.xml.rels to include the comments relationship
 * 
 * @param {string} packageOxml - The complete pkg:package from getOoxml()
 * @param {string} commentsXml - The comments.xml content
 * @returns {string} Complete pkg:package with comments part added
 */
export function injectCommentsIntoPackage(packageOxml, commentsXml) {
    const parser = createParser();
    const serializer = createSerializer();

    const pkgDoc = parser.parseFromString(packageOxml, 'text/xml');

    // Check for parse errors
    const parseError = pkgDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        logError('[CommentEngine] Failed to parse package:', parseError.textContent);
        return packageOxml;
    }

    const pkgPackage = pkgDoc.documentElement;
    const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
    const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

    // 1. Add the comments.xml part
    const commentsPart = pkgDoc.createElementNS(PKG_NS, 'pkg:part');
    commentsPart.setAttribute('pkg:name', '/word/comments.xml');
    commentsPart.setAttribute('pkg:contentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');

    const commentsXmlData = pkgDoc.createElementNS(PKG_NS, 'pkg:xmlData');

    // Parse the comments XML and import it
    const commentsDoc = parser.parseFromString(commentsXml, 'text/xml');
    const commentsNode = pkgDoc.importNode(commentsDoc.documentElement, true);
    commentsXmlData.appendChild(commentsNode);
    commentsPart.appendChild(commentsXmlData);
    pkgPackage.appendChild(commentsPart);

    // 2. Update document.xml.rels to include comments relationship
    const parts = Array.from(pkgPackage.getElementsByTagNameNS(PKG_NS, 'part'));
    const docRelsPart = parts.find(p => {
        const name = p.getAttribute('pkg:name');
        return name === '/word/_rels/document.xml.rels';
    });

    if (docRelsPart) {
        // Find the Relationships element inside
        const xmlDataNodes = docRelsPart.getElementsByTagNameNS(PKG_NS, 'xmlData');
        if (xmlDataNodes.length > 0) {
            const xmlData = xmlDataNodes[0];
            const relsNodes = xmlData.getElementsByTagNameNS(RELS_NS, 'Relationships');
            if (relsNodes.length > 0) {
                const relationships = relsNodes[0];

                // Check if comments relationship already exists
                const existingRels = Array.from(relationships.getElementsByTagNameNS(RELS_NS, 'Relationship'));
                const hasComments = existingRels.some(r =>
                    r.getAttribute('Type')?.includes('comments')
                );

                if (!hasComments) {
                    // Generate a unique rId by finding the highest current rId
                    let maxId = 0;
                    existingRels.forEach(r => {
                        const id = r.getAttribute('Id');
                        const num = parseInt(id?.replace('rId', '') || '0', 10);
                        if (num > maxId) maxId = num;
                    });

                    const newRel = pkgDoc.createElementNS(RELS_NS, 'Relationship');
                    newRel.setAttribute('Id', `rId${maxId + 1}`);
                    newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments');
                    newRel.setAttribute('Target', 'comments.xml');
                    relationships.appendChild(newRel);
                }
            }
        }
    } else {
        // No document.xml.rels exists, create it
        const docRelsPart = pkgDoc.createElementNS(PKG_NS, 'pkg:part');
        docRelsPart.setAttribute('pkg:name', '/word/_rels/document.xml.rels');
        docRelsPart.setAttribute('pkg:contentType', 'application/vnd.openxmlformats-package.relationships+xml');

        const relsXmlData = pkgDoc.createElementNS(PKG_NS, 'pkg:xmlData');
        const relsElem = pkgDoc.createElementNS(RELS_NS, 'Relationships');
        const commentsRel = pkgDoc.createElementNS(RELS_NS, 'Relationship');
        commentsRel.setAttribute('Id', 'rId1');
        commentsRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments');
        commentsRel.setAttribute('Target', 'comments.xml');
        relsElem.appendChild(commentsRel);
        relsXmlData.appendChild(relsElem);
        docRelsPart.appendChild(relsXmlData);
        pkgPackage.appendChild(docRelsPart);
    }

    return serializer.serializeToString(pkgDoc);
}

/**
 * @deprecated Use injectCommentsIntoPackage instead
 * Wraps document OOXML with comments part for insertOoxml.
 * 
 * @param {string} documentXml - The document body XML (w:body content)
 * @param {string} commentsXml - The comments.xml content
 * @returns {string} Complete pkg:package for insertOoxml
 */
export function wrapWithCommentsPart(documentXml, commentsXml) {
    // Build the complete package with comments part
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml">
    <pkg:xmlData>
      ${commentsXml}
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      ${documentXml}
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

