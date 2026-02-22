function clampWordListLevel(level) {
  return Math.max(0, Math.min(8, level));
}

function toInteger(value, fallback = 0) {
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

/**
 * Resolves the target Word list level for insert_list_item calls.
 *
 * The insert_list_item contract supports relative indentation only:
 * -1 = shallower, 0 = same level, 1 = deeper.
 * Values outside this range are clamped to keep insertion stable.
 *
 * @param {number} baseIlvl
 * @param {number} indentLevel
 * @returns {{ baseIlvl: number, normalizedIndent: number, appliedIndent: number, newIlvl: number }}
 */
export function resolveInsertListItemLevel(baseIlvl, indentLevel) {
  const safeBaseIlvl = clampWordListLevel(toInteger(baseIlvl, 0));
  const normalizedIndent = toInteger(indentLevel, 0);
  const appliedIndent = Math.max(-1, Math.min(1, normalizedIndent));
  const newIlvl = clampWordListLevel(safeBaseIlvl + appliedIndent);

  return {
    baseIlvl: safeBaseIlvl,
    normalizedIndent,
    appliedIndent,
    newIlvl
  };
}
