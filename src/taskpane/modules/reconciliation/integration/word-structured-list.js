/**
 * Legacy direct structured list insertion for Word runtime.
 *
 * Kept as a fallback path when reconciliation-driven list insertion fails.
 */

import { insertOoxmlWithRangeFallback } from './word-ooxml.js';
import { inferNumberingStyleFromMarker } from '../orchestration/list-markdown.js';

/**
 * Applies structured list items via direct OOXML insertion against Word paragraph APIs.
 *
 * @param {Word.RequestContext} context - Word request context
 * @param {Word.Paragraph} targetParagraph - Target paragraph anchor
 * @param {{ type?: string, items?: Array<{ type: string, marker?: string, text?: string, level?: number }> }} parsedListData - Parsed list model
 * @returns {Promise<void>}
 */
export async function applyStructuredListDirectOoxml(context, targetParagraph, parsedListData) {
    const listItems = (parsedListData?.items || []).filter((item) => item && (item.type === 'numbered' || item.type === 'bullet'));
    if (listItems.length === 0) {
        throw new Error('No list items parsed for structured list conversion.');
    }

    const listType = parsedListData.type === 'bullet' ? 'bullet' : 'numbered';
    const numberingStyle = listType === 'numbered'
        ? inferNumberingStyleFromMarker(listItems[0].marker || '')
        : 'decimal';

    const numFmtMap = {
        decimal: 'decimal',
        lowerAlpha: 'lowerLetter',
        upperAlpha: 'upperLetter',
        lowerRoman: 'lowerRoman',
        upperRoman: 'upperRoman'
    };
    const numFmt = numFmtMap[numberingStyle] || 'decimal';

    let templateNumId = listType === 'bullet' ? '1' : '2';
    const useCustomNumbering = listType === 'numbered' && numFmt !== 'decimal';

    const firstItem = listItems[0];
    const firstTextEscaped = escapeXmlText(firstItem.text || '');
    const firstIlvl = Math.max(0, Math.min(8, firstItem.level || 0));

    let numberingPart = '';
    if (useCustomNumbering) {
        templateNumId = '100';
        numberingPart = `
            <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
              <pkg:xmlData>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
                </Relationships>
              </pkg:xmlData>
            </pkg:part>
            <pkg:part pkg:name="/word/numbering.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml">
              <pkg:xmlData>
                <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:abstractNum w:abstractNumId="100">
                    <w:multiLevelType w:val="multilevel"/>
                    <w:lvl w:ilvl="0">
                      <w:start w:val="1"/>
                      <w:numFmt w:val="${numFmt}"/>
                      <w:lvlText w:val="%1."/>
                      <w:lvlJc w:val="left"/>
                      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
                    </w:lvl>
                  </w:abstractNum>
                  <w:num w:numId="100">
                    <w:abstractNumId w:val="100"/>
                  </w:num>
                </w:numbering>
              </pkg:xmlData>
            </pkg:part>`;
    }

    const firstParagraphOoxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
            <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
              <pkg:xmlData>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
              </pkg:xmlData>
            </pkg:part>${numberingPart}
            <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
              <pkg:xmlData>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val="ListParagraph"/>
                        <w:numPr>
                          <w:ilvl w:val="${firstIlvl}"/>
                          <w:numId w:val="${templateNumId}"/>
                        </w:numPr>
                      </w:pPr>
                      <w:r>
                        <w:t xml:space="preserve">${firstTextEscaped}</w:t>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
              </pkg:xmlData>
            </pkg:part>
          </pkg:package>`;

    await insertOoxmlWithRangeFallback(targetParagraph, firstParagraphOoxml, 'Replace', context, 'StructuredListDirect/First');

    let resolvedNumId = templateNumId;
    try {
        const firstParaOoxmlResult = targetParagraph.getOoxml();
        await context.sync();
        const numIdMatch = firstParaOoxmlResult.value.match(/<[\w:]*?numId\s+[\w:]*?val="(\d+)"/i);
        if (numIdMatch && numIdMatch[1]) {
            resolvedNumId = numIdMatch[1];
        }
    } catch (resolveError) {
        console.warn('[StructuredListDirect] Could not resolve numId from first inserted item:', resolveError.message);
    }

    let anchorParagraph = targetParagraph;
    for (let i = 1; i < listItems.length; i++) {
        const item = listItems[i];
        const ilvl = Math.max(0, Math.min(8, item.level || 0));
        const textEscaped = escapeXmlText(item.text || '');

        const appendParagraphOoxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
            <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
              <pkg:xmlData>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                </Relationships>
              </pkg:xmlData>
            </pkg:part>
            <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
              <pkg:xmlData>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p>
                      <w:pPr>
                        <w:pStyle w:val="ListParagraph"/>
                        <w:numPr>
                          <w:ilvl w:val="${ilvl}"/>
                          <w:numId w:val="${resolvedNumId}"/>
                        </w:numPr>
                      </w:pPr>
                      <w:r>
                        <w:t xml:space="preserve">${textEscaped}</w:t>
                      </w:r>
                    </w:p>
                  </w:body>
                </w:document>
              </pkg:xmlData>
            </pkg:part>
          </pkg:package>`;

        await insertOoxmlWithRangeFallback(anchorParagraph, appendParagraphOoxml, 'After', context, 'StructuredListDirect/Append');

        const nextParagraph = anchorParagraph.getNextOrNullObject();
        nextParagraph.load('isNullObject');
        await context.sync();
        if (!nextParagraph.isNullObject) {
            anchorParagraph = nextParagraph;
        }
    }
}

function escapeXmlText(text) {
    return String(text || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}
