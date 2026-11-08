/** Parsing for Glance check answers. Model output is untrusted, so only known values pass. */
const STATUSES = new Set(['green', 'yellow', 'red']);
const MAX_SUMMARY_CHARS = 200;

/** All non-thought text of the first candidate; search grounding can split it across parts. */
export function glanceResponseText(result) {
  const parts = result?.candidates?.[0]?.content?.parts;
  if (!Array.isArray(parts)) throw new Error('Gemini returned no Glance answer.');
  const text = parts.filter(part => typeof part?.text === 'string' && !part.thought).map(part => part.text).join('');
  if (!text.trim()) throw new Error('Gemini returned an empty Glance answer.');
  return text;
}

/**
 * Read `{ "<id>": { status, summary } }` from the answer for the given ids.
 * Unknown statuses show as gray; missing ids are left out.
 */
export function parseGlanceResults(text, ids) {
  const match = String(text).match(/\{[\s\S]*\}/);
  if (!match) throw new Error('The Glance answer did not contain JSON.');
  const json = JSON.parse(match[0]);
  const results = {};
  for (const id of ids) {
    const entry = json?.[id];
    if (!entry || typeof entry !== 'object') continue;
    results[id] = {
      status: STATUSES.has(entry.status) ? entry.status : 'gray',
      summary: typeof entry.summary === 'string' ? entry.summary.slice(0, MAX_SUMMARY_CHARS) : '',
    };
  }
  return results;
}
