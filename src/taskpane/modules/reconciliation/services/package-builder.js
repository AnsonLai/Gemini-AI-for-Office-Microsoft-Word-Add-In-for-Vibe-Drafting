/**
 * Shared OOXML package builders.
 *
 * Centralizes `pkg:package` construction used by pipeline/engine/services.
 */

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

const REL_OFFICE_DOCUMENT = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>';
const REL_NUMBERING = '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>';
const REL_COMMENTS = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>';

const DEFAULT_NUMBERING_XML = `
<w:numbering xmlns:w="${NS_W}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`.trim();

function stripXmlDeclaration(xml) {
    if (!xml) return '';
    return xml.replace(/<\?xml[^>]*\?>/g, '');
}

function buildWordDocument(bodyXml, includeRelationshipsNamespace = true) {
    const rNs = includeRelationshipsNamespace ? ` xmlns:r="${NS_R}"` : '';
    return `<w:document xmlns:w="${NS_W}"${rNs}><w:body>${bodyXml}</w:body></w:document>`;
}

function buildCommentsPart(commentsXml) {
    return `
  <pkg:part pkg:name="/word/comments.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml">
    <pkg:xmlData>
      ${commentsXml}
    </pkg:xmlData>
  </pkg:part>`;
}

function buildNumberingPart(numberingXml) {
    return `
  <pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">
    <pkg:xmlData>
      ${stripXmlDeclaration(numberingXml)}
    </pkg:xmlData>
  </pkg:part>`;
}

function buildPackage(documentXml, documentRelationshipsXml = '', extraPartsXml = '') {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="${NS_PKG}">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="${NS_REL}">
        ${REL_OFFICE_DOCUMENT}
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="${NS_REL}">
        ${documentRelationshipsXml}
      </Relationships>
    </pkg:xmlData>
  </pkg:part>${extraPartsXml}
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      ${documentXml}
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}

/**
 * Builds a package for paragraph/document-fragment insertion.
 *
 * @param {string} paragraphXml - Paragraph OOXML content
 * @param {Object} [options={}] - Packaging options
 * @param {boolean} [options.includeNumbering=false] - Include numbering part + relationship
 * @param {string|null} [options.numberingXml=null] - Custom numbering XML
 * @param {boolean} [options.appendTrailingParagraph=true] - Append trailing blank paragraph
 * @returns {string}
 */
export function buildDocumentFragmentPackage(paragraphXml, options = {}) {
    const {
        includeNumbering = false,
        numberingXml = null,
        appendTrailingParagraph = true
    } = options;

    const bodyXml = appendTrailingParagraph
        ? `${paragraphXml}<w:p><w:pPr></w:pPr></w:p>`
        : paragraphXml;

    const documentXml = buildWordDocument(bodyXml, true);

    let relationshipsXml = '';
    let extraPartsXml = '';

    if (numberingXml) {
        relationshipsXml = REL_NUMBERING;
        extraPartsXml += buildNumberingPart(numberingXml);
    } else if (includeNumbering) {
        relationshipsXml = REL_NUMBERING;
        extraPartsXml += buildNumberingPart(DEFAULT_NUMBERING_XML);
    }

    return buildPackage(documentXml, relationshipsXml, extraPartsXml);
}

/**
 * Builds a minimal package containing only paragraph XML in `word/document.xml`.
 *
 * @param {string} paragraphXml - Paragraph OOXML content
 * @returns {string}
 */
export function buildParagraphOnlyPackage(paragraphXml) {
    const documentXml = buildWordDocument(paragraphXml, true);
    return buildPackage(documentXml);
}

/**
 * Builds a minimal paragraph package with comments relationship + part.
 *
 * @param {string} paragraphXml - Paragraph OOXML content
 * @param {string} commentsXml - Comments part content
 * @returns {string}
 */
export function buildParagraphCommentsPackage(paragraphXml, commentsXml) {
    const documentXml = buildWordDocument(paragraphXml, false);
    return buildPackage(documentXml, REL_COMMENTS, buildCommentsPart(commentsXml));
}

/**
 * Builds a package around caller-provided document XML, adding comments part/relationship.
 *
 * @param {string} documentXml - Document XML payload for `word/document.xml`
 * @param {string} commentsXml - Comments part content
 * @returns {string}
 */
export function buildDocumentCommentsPackage(documentXml, commentsXml) {
    return buildPackage(documentXml, REL_COMMENTS, buildCommentsPart(commentsXml));
}
