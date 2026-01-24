/**
 * OOXML Reconciliation Pipeline - Serialization
 * 
 * Converts patched run model back to OOXML with track changes.
 */

import { NS_W, RunKind, escapeXml, getNextRevisionId } from './types.js';
import { getApplicableFormatHints } from './markdown-processor.js';

/**
 * Serializes a patched run model to OOXML.
 * 
 * @param {import('./types.js').RunEntry[]} patchedModel - The patched run model
 * @param {Element|null} pPr - Paragraph properties element
 * @param {import('./types.js').FormatHint[]} [formatHints=[]] - Format hints
 * @param {Object} [options={}] - Serialization options
 * @param {string} [options.author='AI'] - Author for track changes
 * @returns {string} OOXML paragraph string (WITHOUT namespace - added by wrapper)
 */
export function serializeToOoxml(patchedModel, pPr, formatHints = [], options = {}) {
    const { author = 'Gemini AI' } = options;
    const paragraphs = [];
    let currentPPrXml = '';
    let currentRuns = [];

    // Helper to flush accumulated runs into a paragraph
    function flushParagraph() {
        if (currentRuns.length > 0 || paragraphs.length === 0) {
            // Build paragraph properties - handle both string and DOM element
            let pPrContent = '';
            if (currentPPrXml) {
                pPrContent = currentPPrXml.replace(/\s*xmlns:[^=]+="[^"]*"/g, '');
            } else if (pPr) {
                // Fallback to legacy pPr if no PARAGRAPH_START was seen
                if (typeof pPr === 'string') {
                    pPrContent = pPr;
                } else {
                    pPrContent = new XMLSerializer().serializeToString(pPr);
                }
                pPrContent = pPrContent.replace(/\s*xmlns:[^=]+="[^"]*"/g, '');
            }
            paragraphs.push(`<w:p>${pPrContent}${currentRuns.join('')}</w:p>`);
            currentRuns = [];
        }
    }

    for (const item of patchedModel) {
        switch (item.kind) {
            case RunKind.PARAGRAPH_START:
                // Flush previous paragraph before starting a new one
                if (currentRuns.length > 0 || paragraphs.length > 0) {
                    flushParagraph();
                }
                currentPPrXml = item.pPrXml || '';
                break;

            case RunKind.TEXT:
                currentRuns.push(buildRunXmlWithHints(item, formatHints));
                break;

            case RunKind.DELETION:
                currentRuns.push(buildDeletionXml(item, author));
                break;

            case RunKind.INSERTION:
                currentRuns.push(buildInsertionXml(item, formatHints, author));
                break;

            case RunKind.BOOKMARK:
            case RunKind.HYPERLINK:
                // Pass through original XML - but strip any namespace declarations
                if (item.nodeXml) {
                    currentRuns.push(item.nodeXml.replace(/\s*xmlns:[^=]+="[^"]*"/g, ''));
                }
                break;

            case RunKind.CONTAINER_START:
                if (item.containerKind === 'sdt') {
                    currentRuns.push(`<w:sdt>${item.propertiesXml}<w:sdtContent>`);
                } else if (item.containerKind === 'smartTag') {
                    currentRuns.push(`<w:smartTag ${item.propertiesXml}>`);
                } else if (item.containerKind === 'hyperlink') {
                    const props = JSON.parse(item.propertiesXml);
                    const rIdAttr = props.rId ? ` r:id="${props.rId}"` : '';
                    const anchorAttr = props.anchor ? ` w:anchor="${props.anchor}"` : '';
                    currentRuns.push(`<w:hyperlink${rIdAttr}${anchorAttr}>`);
                }
                break;

            case RunKind.CONTAINER_END:
                if (item.containerKind === 'sdt') {
                    currentRuns.push(`</w:sdtContent></w:sdt>`);
                } else if (item.containerKind === 'smartTag') {
                    currentRuns.push(`</w:smartTag>`);
                } else if (item.containerKind === 'hyperlink') {
                    currentRuns.push(`</w:hyperlink>`);
                }
                break;

            default:
                console.warn('Unknown run kind:', item.kind);
        }
    }

    // Flush final paragraph
    flushParagraph();

    // Return all paragraphs WITHOUT namespace - wrapper will add it
    return paragraphs.join('');
}

/**
 * Builds a run XML element, applying format hints if applicable.
 * 
 * @param {import('./types.js').RunEntry} item - Run entry
 * @param {import('./types.js').FormatHint[]} formatHints - Format hints
 * @returns {string}
 */
function buildRunXmlWithHints(item, formatHints) {
    const applicableHints = getApplicableFormatHints(formatHints, item.startOffset, item.endOffset);

    if (applicableHints.length === 0) {
        // No formatting changes - use original rPr (strip namespace)
        const cleanRPr = item.rPrXml ? item.rPrXml.replace(/\s*xmlns:[^=]+="[^"]*"/g, '') : '';
        return buildSimpleRun(item.text, cleanRPr);
    }

    // Split the run text at format boundaries and apply hints
    const runs = [];
    let pos = 0;
    const text = item.text;
    const baseOffset = item.startOffset;
    const cleanRPr = item.rPrXml ? item.rPrXml.replace(/\s*xmlns:[^=]+="[^"]*"/g, '') : '';

    for (const hint of applicableHints) {
        const localStart = Math.max(0, hint.start - baseOffset);
        const localEnd = Math.min(text.length, hint.end - baseOffset);

        // Text before the hint
        if (localStart > pos) {
            runs.push(buildSimpleRun(text.slice(pos, localStart), cleanRPr));
        }

        // Formatted text
        const formattedRPr = injectFormatting(cleanRPr, hint.format);
        runs.push(buildSimpleRun(text.slice(localStart, localEnd), formattedRPr));
        pos = localEnd;
    }

    // Remaining text after last hint
    if (pos < text.length) {
        runs.push(buildSimpleRun(text.slice(pos), cleanRPr));
    }

    return runs.join('');
}

/**
 * Builds a simple w:r element.
 * 
 * @param {string} text - Text content
 * @param {string} rPrXml - Run properties XML
 * @returns {string}
 */
function buildSimpleRun(text, rPrXml) {
    if (!text) return '';
    const rPr = rPrXml || '';
    return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

/**
 * Builds a deletion (w:del) element.
 * 
 * @param {import('./types.js').RunEntry} item - Deletion entry
 * @param {string} author - Author name
 * @returns {string}
 */
function buildDeletionXml(item, author) {
    const revId = getNextRevisionId();
    const date = new Date().toISOString();
    const rPr = item.rPrXml ? item.rPrXml.replace(/\s*xmlns:[^=]+="[^"]*"/g, '') : '';

    return `<w:del w:id="${revId}" w:author="${escapeXml(author)}" w:date="${date}">` +
        `<w:r>${rPr}<w:delText xml:space="preserve">${escapeXml(item.text)}</w:delText></w:r>` +
        `</w:del>`;
}

/**
 * Builds an insertion (w:ins) element.
 * 
 * @param {import('./types.js').RunEntry} item - Insertion entry
 * @param {import('./types.js').FormatHint[]} formatHints - Format hints
 * @param {string} author - Author name
 * @returns {string}
 */
function buildInsertionXml(item, formatHints, author) {
    const revId = getNextRevisionId();
    const date = new Date().toISOString();

    // Build the inner run content with format hints
    const applicableHints = getApplicableFormatHints(formatHints, item.startOffset, item.endOffset);
    let innerContent = '';
    const cleanRPr = item.rPrXml ? item.rPrXml.replace(/\s*xmlns:[^=]+="[^"]*"/g, '') : '';

    if (applicableHints.length === 0) {
        innerContent = buildSimpleRun(item.text, cleanRPr);
    } else {
        // Apply format hints
        let pos = 0;
        const text = item.text;
        const baseOffset = item.startOffset;

        for (const hint of applicableHints) {
            const localStart = Math.max(0, hint.start - baseOffset);
            const localEnd = Math.min(text.length, hint.end - baseOffset);

            if (localStart > pos) {
                innerContent += buildSimpleRun(text.slice(pos, localStart), cleanRPr);
            }

            const formattedRPr = injectFormatting(cleanRPr, hint.format);
            innerContent += buildSimpleRun(text.slice(localStart, localEnd), formattedRPr);
            pos = localEnd;
        }

        if (pos < text.length) {
            innerContent += buildSimpleRun(text.slice(pos), cleanRPr);
        }
    }

    return `<w:ins w:id="${revId}" w:author="${escapeXml(author)}" w:date="${date}">` +
        innerContent +
        `</w:ins>`;
}

/**
 * Injects formatting into run properties XML.
 * 
 * @param {string} baseRPrXml - Base run properties
 * @param {Object} format - Format flags (bold, italic, underline, strikethrough)
 * @returns {string}
 */
function injectFormatting(baseRPrXml, format) {
    if (!format || Object.keys(format).length === 0) {
        return baseRPrXml;
    }

    // Extract existing content from rPr
    let content = '';
    if (baseRPrXml) {
        content = baseRPrXml.replace(/<\/?w:rPr[^>]*>/g, '');
    }

    // Add new formatting elements
    if (format.bold && !content.includes('<w:b')) {
        content = '<w:b/>' + content;
    }
    if (format.italic && !content.includes('<w:i')) {
        content = '<w:i/>' + content;
    }
    if (format.underline && !content.includes('<w:u')) {
        content = '<w:u w:val="single"/>' + content;
    }
    if (format.strikethrough && !content.includes('<w:strike')) {
        content = '<w:strike/>' + content;
    }

    return `<w:rPr>${content}</w:rPr>`;
}

/**
 * Wraps OOXML paragraph content for Word's insertOoxml API.
 * Must include both the document part AND the relationships part.
 * 
 * @param {string} paragraphXml - The paragraph XML (without namespace declarations)
 * @param {Object} [options={}] - Options
 * @param {boolean} [options.includeNumbering=false] - Whether to include numbering definitions
 * @param {string} [options.numberingXml] - Custom numbering XML (w:numbering)
 * @returns {string} Complete OOXML package for insertOoxml
 */
export function wrapInDocumentFragment(paragraphXml, options = {}) {
    const { includeNumbering = false, numberingXml = null } = options;

    // Build document relationships - include numbering if needed
    let docRels = '';
    if (includeNumbering || numberingXml) {
        docRels = '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>';
    }

    // Build numbering.xml part if needed
    let numberingPart = '';
    if (numberingXml) {
        // Use custom provided numbering XML
        numberingPart = `
  <pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">
    <pkg:xmlData>
      ${numberingXml.replace(/<\?xml[^>]*\?>/g, '')}
    </pkg:xmlData>
  </pkg:part>`;
    } else if (includeNumbering) {
        // Fallback to default numbering definition
        numberingPart = `
  <pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">
    <pkg:xmlData>
      <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <!-- Bullet list definition -->
        <w:abstractNum w:abstractNumId="0">
          <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
        </w:abstractNum>
        <!-- Numbered list definition -->
        <w:abstractNum w:abstractNumId="1">
          <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
        </w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
        <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
      </w:numbering>
    </pkg:xmlData>
  </pkg:part>`;
    }

    // Word's insertOoxml requires a complete package with relationships
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
        ${docRels}
      </Relationships>
    </pkg:xmlData>
  </pkg:part>${numberingPart}
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          ${paragraphXml}
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
}


