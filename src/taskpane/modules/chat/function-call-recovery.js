/**
 * Returns the names from the current tool list supplied by the chat loop.
 */
function getToolNames(toolNames) {
  if (!Array.isArray(toolNames) && !(toolNames instanceof Set)) return [];
  const entries = [...toolNames];
  const names = [];
  for (const entry of entries) {
    const trimmed = typeof entry === 'string' ? entry.trim() : '';
    if (trimmed && !names.includes(trimmed)) names.push(trimmed);
  }
  return names;
}

function tryParseArgs(rawArgs) {
  if (typeof rawArgs !== 'string') return null;
  const trimmed = rawArgs.trim();
  if (!trimmed) return {};

  try {
    return JSON.parse(trimmed);
  } catch (_) {
    // Continue with the same tolerant normalization used by the taskpane.
  }

  try {
    const normalized = trimmed
      .replace(/^\(\s*/, '')
      .replace(/\s*\)\s*$/, '')
      .replace(/,\s*([}\]])/g, '$1')
      .replace(/([{,]\s*)([A-Za-z_][A-Za-z0-9_]*)\s*:/g, '$1"$2":')
      .replace(/'([^'\\]*(?:\\.[^'\\]*)*)'/g, (_, value) => `"${value.replace(/"/g, '\\"')}"`);
    return JSON.parse(normalized);
  } catch (_) {
    return null;
  }
}

function parseMalformedEditListArgs(rawArgs) {
  if (typeof rawArgs !== 'string' || !rawArgs) return null;

  const parseIntField = (fieldName) => {
    const match = rawArgs.match(new RegExp(`${fieldName}\\s*:\\s*(\\d+)`, 'i'));
    return match ? parseInt(match[1], 10) : null;
  };

  const parseStringField = (fieldName) => {
    const match = rawArgs.match(new RegExp(`${fieldName}\\s*:\\s*([^,}\\]]+)`, 'i'));
    return match ? String(match[1]).trim().replace(/^["']|["']$/g, '') : null;
  };

  const startParagraphIndex = parseIntField('startParagraphIndex');
  const endParagraphIndex = parseIntField('endParagraphIndex');
  const listTypeRaw = parseStringField('listType');
  const numberingStyle = parseStringField('numberingStyle');

  if (!startParagraphIndex || !endParagraphIndex) return null;

  const listType = (listTypeRaw || 'numbered').toLowerCase();
  const normalizedListType = listType === 'bullet' ? 'bullet' : 'numbered';

  let newItems = [];
  const itemsMatch = rawArgs.match(/newItems\s*:\s*\[([\s\S]*?)\](?=\s*,\s*[A-Za-z_][A-Za-z0-9_]*\s*:|\s*$)/i);
  if (itemsMatch && itemsMatch[1]) {
    const itemsRaw = itemsMatch[1]
      .replace(/\r?\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    if (itemsRaw) {
      // Common malformed pattern: unquoted sentence-like items separated by " , "
      const sentenceSplit = itemsRaw
        .split(/(?<=[.;!?])\s*,\s*(?=[A-Z0-9(“"'])/)
        .map((item) => item.trim().replace(/^["']|["']$/g, ''))
        .filter(Boolean);

      if (sentenceSplit.length > 0) {
        newItems = sentenceSplit;
      } else {
        // Fallback for simpler malformed arrays.
        newItems = itemsRaw
          .split(/\s*,\s*/)
          .map((item) => item.trim().replace(/^["']|["']$/g, ''))
          .filter(Boolean);
      }
    }
  }

  const parsed = {
    startParagraphIndex,
    endParagraphIndex,
    listType: normalizedListType
  };

  if (numberingStyle) parsed.numberingStyle = numberingStyle;
  if (newItems.length > 0) parsed.newItems = newItems;
  return parsed;
}

function parseMalformedConvertHeadersArgs(rawArgs) {
  if (typeof rawArgs !== 'string' || !rawArgs) return null;
  const indicesMatch = rawArgs.match(/paragraphIndices\s*:\s*\[([^\]]*)\]/i);
  if (!indicesMatch || !indicesMatch[1]) return null;

  const paragraphIndices = indicesMatch[1]
    .split(/\s*,\s*/)
    .map((value) => parseInt(value.trim(), 10))
    .filter((value) => Number.isFinite(value));

  if (paragraphIndices.length === 0) return null;

  const numberingFormatMatch = rawArgs.match(/numberingFormat\s*:\s*([^,}\]]+)/i);
  const numberingFormat = numberingFormatMatch
    ? String(numberingFormatMatch[1]).trim().replace(/^["']|["']$/g, '')
    : undefined;

  return numberingFormat
    ? { paragraphIndices, numberingFormat }
    : { paragraphIndices };
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Recover a tool call from Gemini's malformed-call diagnostic text.
 *
 * Only names exposed by the caller's current tool list/registry are scanned.
 * Returns `{ name, args }` on recovery and `null` when the message is absent,
 * unparseable, or refers only to unknown tools.
 */
export function recoverFunctionCall(finishMessage, toolNames) {
  if (typeof finishMessage !== 'string' || !finishMessage) return null;
  const knownToolNames = getToolNames(toolNames);
  if (knownToolNames.length === 0) return null;

  for (const toolName of knownToolNames) {
    const regex = new RegExp(`${escapeRegExp(toolName)}\\s*\\(?\\s*(\\{[\\s\\S]*?\\})\\s*\\)?`, 'i');
    const match = finishMessage.match(regex);
    if (!match || !match[1]) continue;

    let parsedArgs = tryParseArgs(match[1]);
    if ((!parsedArgs || typeof parsedArgs !== 'object' || Array.isArray(parsedArgs)) && toolName === 'edit_list') {
      parsedArgs = parseMalformedEditListArgs(match[1]);
    }
    if ((!parsedArgs || typeof parsedArgs !== 'object' || Array.isArray(parsedArgs)) && toolName === 'convert_headers_to_list') {
      parsedArgs = parseMalformedConvertHeadersArgs(match[1]);
    }
    if (!parsedArgs || typeof parsedArgs !== 'object' || Array.isArray(parsedArgs)) continue;

    return { name: toolName, args: parsedArgs };
  }

  // Legacy fallback: retain the old raw redline instruction recovery, but only
  // when apply_redlines is present in the authoritative tool list/registry.
  if (knownToolNames.includes('apply_redlines')) {
    const redlineMatch = finishMessage.match(/apply_redlines\s*\{\s*instruction\s*:\s*(.*)\s*\}/s);
    if (redlineMatch && redlineMatch[1]) {
      return {
        name: 'apply_redlines',
        args: { instruction: redlineMatch[1].trim() }
      };
    }
  }

  return null;
}
