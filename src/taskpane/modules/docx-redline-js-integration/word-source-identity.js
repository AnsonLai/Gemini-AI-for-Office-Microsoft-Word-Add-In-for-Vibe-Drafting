import { computeDocumentPartsRevisionToken, inspectDocumentParts, NS_W } from '@ansonlai/docx-redline-js';
import { parseXml, serializeXml } from '@ansonlai/docx-redline-js/adapters/xml-adapter.js';

const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const VOLATILE_WORD_RSID_ATTRIBUTES = new Set([
  'rsidR', 'rsidRDefault', 'rsidRPr', 'rsidP', 'rsidDel', 'rsidSect', 'rsidTr'
]);

function localName(node) {
  return String(node?.localName || node?.nodeName || '').replace(/^.*:/, '');
}

function wordBody(document) {
  return Array.from(document?.getElementsByTagNameNS?.(NS_W, 'body') || [])[0] || null;
}

function bodyParagraphNodes(document) {
  return Array.from(wordBody(document)?.getElementsByTagNameNS?.(NS_W, 'p') || []);
}

function isInsideTable(paragraph) {
  let current = paragraph?.parentNode;
  while (current) {
    if (current.namespaceURI === NS_W && localName(current) === 'tbl') return true;
    current = current.parentNode;
  }
  return false;
}

function hasListProperties(paragraph) {
  const properties = Array.from(paragraph?.childNodes || []).find(node =>
    node.nodeType === 1 && node.namespaceURI === NS_W && localName(node) === 'pPr'
  );
  return Array.from(properties?.childNodes || []).some(node =>
    node.nodeType === 1 && node.namespaceURI === NS_W && localName(node) === 'numPr'
  );
}

/**
 * Returns the one-based canonical paragraph index for Word's export-only tail,
 * or null unless the host and library counts certify the exact known shape.
 * The final canonical paragraph must be empty and outside tables and lists.
 */
export function detectExportTailIndex(parts, { hostParagraphCount } = {}) {
  if (!Number.isInteger(hostParagraphCount) || hostParagraphCount < 0
    || typeof parts?.documentXml !== 'string') return null;
  let inspection;
  let document;
  try {
    inspection = inspectDocumentParts(parts);
    if (inspection?.status !== 'ok') return null;
    document = parseXml(parts.documentXml);
  } catch {
    return null;
  }
  const paragraphs = inspection.paragraphs || [];
  const last = paragraphs.at(-1);
  const xmlParagraphs = bodyParagraphNodes(document);
  if (xmlParagraphs.length !== hostParagraphCount + 1
    || last?.index !== xmlParagraphs.length
    || !last || last.exactText !== '' || last.inTable === true || last.list != null) return null;
  const xmlTail = xmlParagraphs.at(-1);
  if (!xmlTail || isInsideTable(xmlTail) || hasListProperties(xmlTail)) return null;
  return last.index;
}

function elementsIncludingRoot(document) {
  if (!document?.documentElement) return [];
  return [document.documentElement, ...Array.from(document.documentElement.getElementsByTagName('*'))];
}

function normalizeDocumentXml(xml, exportTailIndex) {
  if (typeof xml !== 'string') return xml;
  const document = parseXml(xml);
  for (const element of elementsIncludingRoot(document)) {
    for (const attribute of Array.from(element.attributes || [])) {
      const name = localName(attribute);
      if (attribute.namespaceURI === NS_W && VOLATILE_WORD_RSID_ATTRIBUTES.has(name)) {
        element.removeAttributeNS(NS_W, name);
      }
    }
  }
  if (exportTailIndex != null) {
    const tail = bodyParagraphNodes(document)[exportTailIndex - 1];
    const paragraphId = Array.from(tail?.attributes || []).find(attribute =>
      attribute.namespaceURI === NS_W14 && localName(attribute) === 'paraId'
    );
    if (paragraphId) tail.removeAttributeNS(NS_W14, 'paraId');
  }
  return serializeXml(document);
}

function normalizeSettingsXml(xml) {
  if (typeof xml !== 'string') return xml;
  const document = parseXml(xml);
  const root = document.documentElement;
  if (root?.namespaceURI === NS_W && localName(root) === 'settings') {
    for (const child of Array.from(root.childNodes || [])) {
      if (child.nodeType === 1 && child.namespaceURI === NS_W && localName(child) === 'rsids') {
        root.removeChild(child);
      }
    }
  }
  return serializeXml(document);
}

function normalizedAdditionalPart(name, payload) {
  const normalizedName = String(name || '').replace(/^\//, '');
  if (normalizedName === 'word/settings.xml') return normalizeSettingsXml(payload);
  if (normalizedName.startsWith('word/') && normalizedName.endsWith('.xml') && typeof payload === 'string') return normalizeDocumentXml(payload, null);
  return payload;
}

function tokenParts(parts, tailIndex) {
  const derived = { ...parts };
  derived.documentXml = normalizeDocumentXml(parts.documentXml, tailIndex);
  for (const name of ['commentsXml', 'commentsExtendedXml', 'commentsIdsXml', 'commentsExtensibleXml', 'numberingXml', 'stylesXml']) {
    if (typeof parts[name] === 'string') derived[name] = normalizeDocumentXml(parts[name], null);
  }
  if (parts.additionalParts && typeof parts.additionalParts === 'object') {
    derived.additionalParts = Object.fromEntries(Object.entries(parts.additionalParts)
      .map(([name, payload]) => [name, normalizedAdditionalPart(name, payload)]));
  }
  if (parts.parts instanceof Map) {
    derived.parts = new Map([...parts.parts.entries()]
      .map(([name, payload]) => [name, normalizedAdditionalPart(name, payload)]));
  }
  return derived;
}

/**
 * Compute a Word-aware identity over a derived copy of canonical parts.
 * Word session IDs and a certified export-only tail ID are omitted only from
 * this token input; the source OOXML returned to inspectors/checkpoints stays
 * byte-for-byte untouched.
 */
export async function computeWordSourceRevisionToken(parts, options = {}) {
  const exportTail = detectExportTailIndex(parts, options);
  return computeDocumentPartsRevisionToken(tokenParts(parts, exportTail));
}
