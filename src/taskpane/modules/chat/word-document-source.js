/** Word transport only. Canonical text/metadata come from the library inspector. */
export function createWordDocumentSourceReader({ loadSupport, getDocumentId, signal, run = callback => Word.run(callback) }) {
  return async function readSource() {
    const abort = () => {
      if (signal?.aborted) throw Object.assign(new Error('Request cancelled before reading document context.'), { name: 'AbortError' });
    };
    abort();
    const [support] = await loadSupport();
    abort();
    const documentId = getDocumentId();
    let sourceOoxml, selectionText;
    await run(async context => {
      const source = context.document.body.getOoxml();
      const selection = context.document.getSelection();
      selection.load('text');
      await context.sync();
      sourceOoxml = source.value;
      selectionText = selection.text || '';
    });
    abort();
    if (documentId !== getDocumentId()) throw Object.assign(new Error('The active document changed during inspection.'), { code: 'DOCUMENT_CHANGED' });
    const { parts, inspection } = support.inspectWordSource(sourceOoxml);
    return { documentId, revisionView: 'accepted', parts, inspection, sourceOoxml, selectionText };
  };
}

/** Same accepted text and library metadata that supply strong edit descriptors. */
export function formatCanonicalDocumentContext(inspection) {
  if (inspection?.status !== 'ok') throw Object.assign(new Error('Canonical document context is unavailable.'), { code: 'CANONICAL_CONTEXT_UNAVAILABLE' });
  return inspection.paragraphs.map(paragraph => {
    const meta = [paragraph.styleId || 'Normal'];
    if (paragraph.list) meta.push(`L${paragraph.list.numId}:${paragraph.list.level}`);
    if (paragraph.table) meta.push(`T:${paragraph.table.rowIndex},${paragraph.table.cellIndex}`);
    if (paragraph.headingLevel) meta.push('§');
    if (paragraph.hasRevisions && !paragraph.exactText) meta.push('deleted');
    return `[P${paragraph.index}|${meta.join('|')}] ${paragraph.exactText}`;
  }).join('\n');
}
