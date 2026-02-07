/**
 * Comment package builders and pkg:part wiring.
 */

import { createParser, createSerializer } from '../adapters/xml-adapter.js';
import { error as logError } from '../adapters/logger.js';
import { buildDocumentCommentsPackage, buildParagraphCommentsPackage } from './package-builder.js';
import { getElementsByTagNS, getXmlParseError } from '../core/xml-query.js';

const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/**
 * Wraps paragraph XML with minimal package structure including comments part.
 *
 * @param {string} paragraphXml - Paragraph XML content
 * @param {string} commentsXml - comments.xml payload
 * @returns {string}
 */
export function wrapParagraphWithComments(paragraphXml, commentsXml) {
    return buildParagraphCommentsPackage(paragraphXml, commentsXml);
}

/**
 * Injects comments.xml and relationship entry into an existing pkg:package.
 *
 * @param {string} packageOxml - Existing package XML
 * @param {string} commentsXml - comments.xml payload
 * @returns {string}
 */
export function injectCommentsIntoPackage(packageOxml, commentsXml) {
    const parser = createParser();
    const serializer = createSerializer();
    const pkgDoc = parser.parseFromString(packageOxml, 'text/xml');

    const parseError = getXmlParseError(pkgDoc);
    if (parseError) {
        logError('[CommentEngine] Failed to parse package:', parseError.textContent);
        return packageOxml;
    }

    const pkgPackage = pkgDoc.documentElement;

    const commentsPart = pkgDoc.createElementNS(PKG_NS, 'pkg:part');
    commentsPart.setAttribute('pkg:name', '/word/comments.xml');
    commentsPart.setAttribute('pkg:contentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml');

    const commentsXmlData = pkgDoc.createElementNS(PKG_NS, 'pkg:xmlData');
    const commentsDoc = parser.parseFromString(commentsXml, 'text/xml');
    commentsXmlData.appendChild(pkgDoc.importNode(commentsDoc.documentElement, true));
    commentsPart.appendChild(commentsXmlData);
    pkgPackage.appendChild(commentsPart);

    const parts = getElementsByTagNS(pkgPackage, PKG_NS, 'part');
    const docRelsPart = parts.find(part => part.getAttribute('pkg:name') === '/word/_rels/document.xml.rels');

    if (docRelsPart) {
        const xmlDataNodes = getElementsByTagNS(docRelsPart, PKG_NS, 'xmlData');
        if (xmlDataNodes.length > 0) {
            const relsNodes = getElementsByTagNS(xmlDataNodes[0], RELS_NS, 'Relationships');
            if (relsNodes.length > 0) {
                const relationships = relsNodes[0];
                const existingRels = getElementsByTagNS(relationships, RELS_NS, 'Relationship');
                const hasCommentsRel = existingRels.some(rel =>
                    rel.getAttribute('Type')?.includes('comments')
                );

                if (!hasCommentsRel) {
                    let maxId = 0;
                    existingRels.forEach(rel => {
                        const id = rel.getAttribute('Id');
                        const idNumber = parseInt(id?.replace('rId', '') || '0', 10);
                        if (idNumber > maxId) maxId = idNumber;
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
        const newDocRelsPart = pkgDoc.createElementNS(PKG_NS, 'pkg:part');
        newDocRelsPart.setAttribute('pkg:name', '/word/_rels/document.xml.rels');
        newDocRelsPart.setAttribute('pkg:contentType', 'application/vnd.openxmlformats-package.relationships+xml');

        const relsXmlData = pkgDoc.createElementNS(PKG_NS, 'pkg:xmlData');
        const relationships = pkgDoc.createElementNS(RELS_NS, 'Relationships');
        const commentsRel = pkgDoc.createElementNS(RELS_NS, 'Relationship');
        commentsRel.setAttribute('Id', 'rId1');
        commentsRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments');
        commentsRel.setAttribute('Target', 'comments.xml');
        relationships.appendChild(commentsRel);
        relsXmlData.appendChild(relationships);
        newDocRelsPart.appendChild(relsXmlData);
        pkgPackage.appendChild(newDocRelsPart);
    }

    return serializer.serializeToString(pkgDoc);
}

/**
 * @deprecated Use injectCommentsIntoPackage instead.
 *
 * @param {string} documentXml - Document OOXML
 * @param {string} commentsXml - comments.xml payload
 * @returns {string}
 */
export function wrapWithCommentsPart(documentXml, commentsXml) {
    return buildDocumentCommentsPackage(documentXml, commentsXml);
}
