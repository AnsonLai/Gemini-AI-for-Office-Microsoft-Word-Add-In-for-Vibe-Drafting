const RECOVERABLE_TOOLS = new Set(['inspect_document', 'apply_document_edits']);

function balancedJsonObject(text, start) {
  if (text[start] !== '{') return null;
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let index = start; index < text.length; index++) {
    const character = text[index];
    if (inString) {
      if (escaped) escaped = false;
      else if (character === '\\') escaped = true;
      else if (character === '"') inString = false;
      continue;
    }
    if (character === '"') inString = true;
    else if (character === '{') depth++;
    else if (character === '}' && --depth === 0) return text.slice(start, index + 1);
  }
  return null;
}

/** Recover only complete JSON argument objects for tools in the live registry. */
export function recoverFunctionCall(finishMessage, toolNames) {
  if (typeof finishMessage !== 'string' || !Array.isArray(toolNames) && !(toolNames instanceof Set)) return null;
  const names = [...new Set([...toolNames].filter(name => RECOVERABLE_TOOLS.has(name)))];
  for (const name of names) {
    const match = new RegExp(`\\b${name}\\b`, 'i').exec(finishMessage);
    if (!match) continue;
    const argumentOffset = match.index + match[0].length;
    const opening = /^\s*(\()?\s*/.exec(finishMessage.slice(argumentOffset));
    const start = argumentOffset + opening[0].length;
    if (finishMessage[start] !== '{') continue;
    const json = balancedJsonObject(finishMessage, start);
    if (!json) continue;
    if (opening[1] && finishMessage.slice(start + json.length).trimStart()[0] !== ')') continue;
    try {
      const args = JSON.parse(json);
      if (args && typeof args === 'object' && !Array.isArray(args)) return { name, args };
    } catch {
      // A malformed diagnostic is recoverable only when the complete JSON parses.
    }
  }
  return null;
}
