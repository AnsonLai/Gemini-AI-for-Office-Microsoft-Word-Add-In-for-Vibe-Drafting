/** Word transport only. Canonical text/metadata come from the library inspector. */
export function queueHostParagraphCount(body) {
  const paragraphs = body.paragraphs;
  if (!paragraphs?.load) return () => undefined;
  paragraphs.load('$none');
  return () => paragraphs.items.length;
}

export function createWordDocumentSourceReader({ loadSupport, getDocumentId, signal, run = callback => Word.run(callback) }) {
  return async function readSource() {
    const abort = () => {
      if (signal?.aborted) throw Object.assign(new Error('Request cancelled before reading document context.'), { name: 'AbortError' });
    };
    abort();
    const [support] = await loadSupport();
    abort();
    const documentId = getDocumentId();
    let sourceOoxml, selectionText, hostParagraphCount;
    await run(async context => {
      const body = context.document.body;
      const source = body.getOoxml();
      const selection = context.document.getSelection();
      selection.load('text');
      const readCount = queueHostParagraphCount(body);
      await context.sync();
      sourceOoxml = source.value;
      selectionText = selection.text || '';
      const count = readCount();
      if (Number.isInteger(count) && count >= 0) hostParagraphCount = count;
    });
    abort();
    if (documentId !== getDocumentId()) throw Object.assign(new Error('The active document changed during inspection.'), { code: 'DOCUMENT_CHANGED' });
    const { parts, inspection } = support.inspectWordSource(sourceOoxml);
    return { documentId, revisionView: 'accepted', parts, inspection, sourceOoxml, selectionText, hostParagraphCount };
  };
}
