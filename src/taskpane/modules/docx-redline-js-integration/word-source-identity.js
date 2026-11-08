import { computeDocumentPartsRevisionToken, inspectDocumentParts, NS_W, parseOoxmlSafe, serializeOoxml as serializeXml } from '@ansonlai/docx-redline-js';

function parseXml(xml) {
  const { doc, error } = parseOoxmlSafe(xml, 'text/xml');
  if (error) throw Object.assign(new Error(error.message), { code: error.code });
  return doc;
}

const NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const NS_W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const NS_W16CID = 'http://schemas.microsoft.com/office/word/2016/wordml/cid';
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

// Word gives footnote and endnote paragraphs new w14:paraId/textId values on every body
// export, even with no edit, so a footnoted document would look changed on each read (KB-16).
// Those IDs never identify edit targets, so the token omits them in these parts only.
const NOTES_PARTS = new Set(['word/footnotes.xml', 'word/endnotes.xml']);

function normalizeNotesXml(xml) {
  if (typeof xml !== 'string') return xml;
  const document = parseXml(normalizeDocumentXml(xml, null));
  for (const element of elementsIncludingRoot(document)) {
    if (element.namespaceURI === NS_W && localName(element) === 'p') {
      element.removeAttributeNS(NS_W14, 'paraId');
      element.removeAttributeNS(NS_W14, 'textId');
    }
  }
  return serializeXml(document);
}

function normalizedAdditionalPart(name, payload) {
  const normalizedName = String(name || '').replace(/^\//, '');
  if (normalizedName === 'word/settings.xml') return normalizeSettingsXml(payload);
  if (NOTES_PARTS.has(normalizedName)) return normalizeNotesXml(payload);
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

// Word's background proofing and layout add these after a write and split runs at
// them, so a document nobody touched would otherwise look changed (see
// logicalRevisionMarkup in scripts/golden/golden-checks.mjs).
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const REVERT_GUARD_NOISE = new Set(['proofErr', 'lastRenderedPageBreak', 'proofState']);

function isSimpleRun(node) {
  if (node?.nodeType !== 1 || node.namespaceURI !== NS_W || localName(node) !== 'r') return false;
  const children = Array.from(node.childNodes || []).filter(child => child.nodeType === 1);
  const text = children.filter(child => child.namespaceURI === NS_W && ['t', 'delText'].includes(localName(child)));
  const properties = children.filter(child => child.namespaceURI === NS_W && localName(child) === 'rPr');
  return text.length === 1 && properties.length <= 1 && text.length + properties.length === children.length;
}

function runParts(run) {
  const children = Array.from(run.childNodes).filter(child => child.nodeType === 1);
  const text = children.find(child => localName(child) !== 'rPr');
  const properties = children.find(child => localName(child) === 'rPr');
  return { text, key: `${localName(text)}|${properties ? serializeXml(properties) : ''}` };
}

function normalizeForRevertGuard(xml) {
  if (typeof xml !== 'string') return xml;
  const document = parseXml(xml);
  const elements = elementsIncludingRoot(document);
  const goBackIds = new Set(elements
    .filter(element => element.namespaceURI === NS_W && localName(element) === 'bookmarkStart' && element.getAttributeNS(NS_W, 'name') === '_GoBack')
    .map(element => element.getAttributeNS(NS_W, 'id')));
  for (const element of elements) {
    if (element.namespaceURI !== NS_W) continue;
    const name = localName(element);
    if (REVERT_GUARD_NOISE.has(name)
      || ((name === 'bookmarkStart' || name === 'bookmarkEnd') && goBackIds.has(element.getAttributeNS(NS_W, 'id')))) {
      element.parentNode?.removeChild(element);
    } else if (name === 'p') {
      element.removeAttributeNS(NS_W14, 'paraId');
      element.removeAttributeNS(NS_W14, 'textId');
    } else if (name === 't' || name === 'delText') {
      element.removeAttributeNS(XML_NS, 'space');
    }
  }
  // A run left with nothing but properties shows nothing.
  for (const run of elementsIncludingRoot(document).filter(element => element.namespaceURI === NS_W && localName(element) === 'r')) {
    if (Array.from(run.childNodes || []).every(child => child.nodeType !== 1 || (child.namespaceURI === NS_W && localName(child) === 'rPr'))) {
      run.parentNode?.removeChild(run);
    }
  }
  // Merge adjacent runs with identical properties, undoing proofing splits.
  for (const parent of elementsIncludingRoot(document)) {
    let previous = null;
    for (const child of Array.from(parent.childNodes || [])) {
      if (child.nodeType === 3 && !child.nodeValue.trim()) continue;
      if (!isSimpleRun(child)) { previous = null; continue; }
      const current = runParts(child);
      if (previous && previous.key === current.key) {
        previous.text.appendChild(document.createTextNode(current.text.textContent));
        parent.removeChild(child);
      } else {
        previous = current;
      }
    }
  }
  return serializeXml(document);
}

function normalizeForRevertGuardSafely(xml) {
  try { return normalizeForRevertGuard(xml); } catch { return xml; }
}

/**
 * Token for the Revert guard: was the document changed after an edit? It also
 * ignores Word's background proofing marks, layout marks, the _GoBack bookmark,
 * paragraph ids and run splits, which Word adds on its own after a write. Text,
 * formatting, revisions and comments still count.
 */
export async function computeWordRevertGuardToken(parts, options = {}) {
  const derived = tokenParts(parts, detectExportTailIndex(parts, options));
  for (const [name, value] of Object.entries(derived)) {
    if (name.endsWith('Xml') && typeof value === 'string') derived[name] = normalizeForRevertGuardSafely(value);
  }
  if (derived.additionalParts && typeof derived.additionalParts === 'object') {
    derived.additionalParts = Object.fromEntries(Object.entries(derived.additionalParts)
      .map(([name, payload]) => [name, /\.xml$/i.test(name) && typeof payload === 'string' ? normalizeForRevertGuardSafely(payload) : payload]));
  }
  return computeDocumentPartsRevisionToken(derived);
}

/**
 * Compare a restored body package. Word regenerates body paragraph/text IDs
 * during Replace, even when content is unchanged. Omit only those IDs from this
 * verification copy; regular freshness tokens and checkpoint XML remain exact.
 * Regenerated comment paragraph links require a separate canonical topology
 * comparison; comment IDs/anchors, revisions and other part content still count.
 */
function normalizeRestoredCommentPart(xml, name, knownParagraphIds) {
  if (typeof xml !== 'string') return xml;
  const document = parseXml(xml);
  for (const element of elementsIncludingRoot(document)) {
    if (name === 'commentsXml' && element.namespaceURI === NS_W && localName(element) === 'p') {
      element.removeAttributeNS(NS_W14, 'paraId');
      element.removeAttributeNS(NS_W14, 'textId');
    } else if (name === 'commentsExtendedXml' && element.namespaceURI === NS_W15 && localName(element) === 'commentEx') {
      const id = element.getAttributeNS(NS_W15, 'paraId');
      const parent = element.getAttributeNS(NS_W15, 'paraIdParent');
      if (!knownParagraphIds.has(id) || (parent && !knownParagraphIds.has(parent))) throw new Error('Checkpoint thread metadata cannot be resolved to inspected comments.');
      element.removeAttributeNS(NS_W15, 'paraId');
      element.removeAttributeNS(NS_W15, 'paraIdParent');
    } else if (name === 'commentsIdsXml' && element.namespaceURI === NS_W16CID && localName(element) === 'commentId') {
      if (!knownParagraphIds.has(element.getAttributeNS(NS_W16CID, 'paraId'))) throw new Error('Checkpoint durable-comment metadata cannot be resolved to inspected comments.');
      element.removeAttributeNS(NS_W16CID, 'paraId');
    }
  }
  return serializeXml(document);
}

export async function computeWordRestorationToken(parts, options = {}) {
  const inspection = inspectDocumentParts(parts);
  if (inspection.status !== 'ok') throw new Error('The checkpoint comment identity could not be inspected.');
  if (inspection.comments?.some(comment => comment.parentParaId && comment.parentCommentId == null)) throw new Error('Checkpoint reply identity cannot be resolved.');
  const knownParagraphIds = new Set((inspection.comments || []).map(comment => comment.paraId).filter(Boolean));
  const commentIdentity = (inspection.comments || []).map(({ paraId, parentParaId, ...comment }) => comment);
  const derived = tokenParts(parts, detectExportTailIndex(parts, options));
  const document = parseXml(derived.documentXml);
  for (const paragraph of bodyParagraphNodes(document)) {
    paragraph.removeAttributeNS(NS_W14, 'paraId');
    paragraph.removeAttributeNS(NS_W14, 'textId');
  }
  derived.documentXml = serializeXml(document);
  const commentParts = { 'word/comments.xml': 'commentsXml', 'word/commentsExtended.xml': 'commentsExtendedXml', 'word/commentsIds.xml': 'commentsIdsXml' };
  for (const name of Object.values(commentParts)) derived[name] = normalizeRestoredCommentPart(derived[name], name, knownParagraphIds);
  if (derived.additionalParts) derived.additionalParts = Object.fromEntries(Object.entries(derived.additionalParts)
    .map(([name, xml]) => [name, commentParts[name.replace(/^\//, '')] ? normalizeRestoredCommentPart(xml, commentParts[name.replace(/^\//, '')], knownParagraphIds) : xml]));
  return { ...await computeDocumentPartsRevisionToken(derived), commentIdentity };
}
