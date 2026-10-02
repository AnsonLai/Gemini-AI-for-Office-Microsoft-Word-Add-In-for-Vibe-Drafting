/**
 * Rebuild the model's anchored document view from the canonical source
 * baseline so every [P#] line shows exactly the text edits are applied to.
 *
 * Word's Paragraph.text still contains tracked-deleted text (and reports soft
 * line breaks as "\v"), while the engine targets the accepted view. A model
 * that copies a `find` from Word text that is pending deletion passes local
 * validation but fails in the engine (PATCH_SOURCE_NOT_FOUND).
 *
 * @param {Array<{index:number, meta:string, text:string}>} paragraphs - extractEnhancedDocumentContext() paragraphs
 * @param {Array<{index:number, exactText:string}>} sourceBaseline - canonical baseline for the same document
 * @returns {string|null} anchored text, or null when the two views do not align one-to-one
 */
export function buildCanonicalContextText(paragraphs, sourceBaseline) {
  if (!Array.isArray(paragraphs) || !Array.isArray(sourceBaseline) || paragraphs.length === 0) {
    return null;
  }
  // Word's body.getOoxml() package carries one extra, empty trailing paragraph
  // that Paragraph collections do not report (observed in desktop Word).
  const trailingArtifact = sourceBaseline.length === paragraphs.length + 1
    && sourceBaseline.at(-1)?.exactText === "";
  if (paragraphs.length !== sourceBaseline.length && !trailingArtifact) {
    return null;
  }
  const aligned = paragraphs.every((paragraph, offset) => (
    paragraph?.index === offset + 1
    && typeof paragraph.meta === "string"
    && sourceBaseline[offset]?.index === offset + 1
    && typeof sourceBaseline[offset].exactText === "string"
  ));
  if (!aligned) return null;

  return paragraphs.map((paragraph, offset) => {
    const canonicalText = sourceBaseline[offset].exactText;
    // Visible in Word but empty in the accepted view: the whole paragraph is a
    // pending tracked deletion and has no editable text.
    const pendingDeletion = canonicalText === "" && String(paragraph.text || "").trim() !== "";
    return `[P${paragraph.index}|${paragraph.meta}${pendingDeletion ? "|deleted" : ""}] ${canonicalText}`;
  }).join("\n");
}
