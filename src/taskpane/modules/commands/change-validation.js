/* global console */
/**
 * change-validation.js
 *
 * Code-level reliability layer for AI-generated redline change sets.
 *
 * This module is intentionally free of any Office.js / Word / DOM dependencies so
 * it can be unit tested in plain Node and reused by the eval harness. It enforces,
 * in code, the rules that are otherwise only stated as prose in the redline prompt.
 *
 * WP1 (this file): content-anchor verification so an edit can never silently land
 * on the wrong paragraph.
 */

/**
 * Normalize a string for anchor comparison: strip a leading [P#] / [P#|meta]
 * marker the model may have copied in, collapse whitespace runs to a single
 * space, and trim.
 *
 * @param {*} value
 * @returns {string}
 */
export function normalizeForAnchor(value) {
  return String(value == null ? "" : value)
    .replace(/^\s*\[P\d+(?:\|[^\]]*)?\]\s*/, "")
    .replace(/\s+/g, " ")
    .trim();
}

function validateLocalizedReplacements(replacements, paragraphText = null) {
  if (!Array.isArray(replacements) || replacements.length === 0) return "invalid_replacements";
  const text = paragraphText == null ? null : String(paragraphText);
  for (const item of replacements) {
    if (!item || typeof item !== "object" || Array.isArray(item)
      || typeof item.find !== "string" || item.find.length === 0
      || typeof item.replace !== "string"
      || (item.occurrence != null && (!Number.isInteger(item.occurrence) || item.occurrence < 1))) {
      return "invalid_replacements";
    }
    if (text != null) {
      const count = text.split(item.find).length - 1;
      if (count === 0 || (item.occurrence != null && item.occurrence > count)) {
        return "replacement_find_not_found";
      }
      if (count > 1 && item.occurrence == null) return "replacement_ambiguous_find";
    }
  }
  return null;
}

/**
 * Parse the anchored document text (lines shaped like `[P#|meta] text`) into an
 * array of paragraph texts where array index 0 corresponds to [P1].
 *
 * Lines that do not begin with a `[P#` header are treated as a continuation of
 * the previous paragraph's text (defensive against rare embedded newlines).
 *
 * @param {string} fullDocumentText
 * @returns {string[]} sparse-safe array; index i holds the text for [P(i+1)]
 */
export function parseAnchoredParagraphs(fullDocumentText) {
  const paragraphs = [];
  if (typeof fullDocumentText !== "string" || fullDocumentText.length === 0) {
    return paragraphs;
  }

  const headerRe = /^\[P(\d+)(?:\|[^\]]*)?\]\s?/;
  const lines = fullDocumentText.split(/\r?\n/);
  let currentIndex = null;

  for (const line of lines) {
    const match = line.match(headerRe);
    if (match) {
      const idx = parseInt(match[1], 10);
      currentIndex = idx;
      paragraphs[idx - 1] = line.slice(match[0].length);
    } else if (currentIndex != null) {
      paragraphs[currentIndex - 1] = `${paragraphs[currentIndex - 1] || ""}\n${line}`;
    }
  }

  return paragraphs;
}

/**
 * Verify that a change's anchorText matches the actual paragraph it claims to
 * target. Tolerates an off-by-one (or off-by-two) index by searching a small
 * neighborhood; only auto-corrects when exactly one neighbor matches.
 *
 * @param {object} change - one change object (paragraphIndex, anchorText, ...)
 * @param {string[]} paragraphTexts - actual paragraph texts, index 0 = [P1]
 * @returns {{ ok: boolean, correctedIndex?: number, reason?: string, actualTextSnippet?: string }}
 */
export function verifyAnchor(change, paragraphTexts) {
  const claimedIndex = change ? change.paragraphIndex : undefined;
  const texts = Array.isArray(paragraphTexts) ? paragraphTexts : [];
  const verifyFinds = (idx) => {
    if (!change || change.replacements == null) return { ok: true };
    if (!Number.isInteger(idx) || idx < 1 || idx > texts.length || texts[idx - 1] == null) {
      return { ok: false, reason: "replacement_find_not_found" };
    }
    const reason = validateLocalizedReplacements(change.replacements, texts[idx - 1]);
    return reason ? { ok: false, reason } : { ok: true };
  };

  // Backwards-compatible: no anchor provided -> skip verification.
  const anchorRaw = change ? change.anchorText : undefined;
  if (anchorRaw == null || String(anchorRaw).trim() === "") {
    console.warn(
      `verifyAnchor: change for P${claimedIndex} has no anchorText; skipping verification.`
    );
    return verifyFinds(claimedIndex);
  }

  const anchor = normalizeForAnchor(anchorRaw);
  if (anchor === "") {
    return verifyFinds(claimedIndex);
  }

  const matchesAt = (idx) => {
    if (!Number.isInteger(idx) || idx < 1 || idx > texts.length) return false;
    const paraText = texts[idx - 1];
    if (paraText == null) return false;
    const normPara = normalizeForAnchor(paraText);
    return normPara.startsWith(anchor) || normPara.includes(anchor);
  };

  // P(n+1) is an append intent. Its anchor describes the current last
  // paragraph, but a matching neighbor must never turn that append into an
  // edit of P(n).
  const appendAllowed =
    change?.operation === "replace_paragraph" ||
    change?.operation === "replace_range" ||
    change?.operation === "edit_paragraph";
  if (claimedIndex === texts.length + 1 && appendAllowed) {
    if (matchesAt(texts.length)) return verifyFinds(claimedIndex);
    const actual = texts.length > 0 ? texts[texts.length - 1] : undefined;
    return {
      ok: false,
      reason: "anchor_mismatch",
      actualTextSnippet: actual != null ? normalizeForAnchor(actual).slice(0, 60) : "",
    };
  }

  // 1. Exact claimed index.
  if (matchesAt(claimedIndex)) {
    return verifyFinds(claimedIndex);
  }

  // 2. Search a ±2 window; only accept an unambiguous single match.
  const neighbors = [];
  for (let delta = 1; delta <= 2; delta++) {
    if (matchesAt(claimedIndex - delta)) neighbors.push(claimedIndex - delta);
    if (matchesAt(claimedIndex + delta)) neighbors.push(claimedIndex + delta);
  }
  if (neighbors.length === 1) {
    const verdict = verifyFinds(neighbors[0]);
    return verdict.ok ? { ok: true, correctedIndex: neighbors[0] } : verdict;
  }

  // 3. No unambiguous match -> reject with what is actually at the claimed index.
  const actual =
    Number.isInteger(claimedIndex) && claimedIndex >= 1 && claimedIndex <= texts.length
      ? texts[claimedIndex - 1]
      : undefined;
  const actualTextSnippet = actual != null ? normalizeForAnchor(actual).slice(0, 60) : "";

  return { ok: false, reason: "anchor_mismatch", actualTextSnippet };
}

// ---------------------------------------------------------------------------
// WP2: mechanical change-set sanitizer. Enforces, in code, the rules that are
// otherwise only stated as prose in the executeRedline prompt, so weaker models
// that ignore the prose still produce a valid change set.
// ---------------------------------------------------------------------------

const VALID_OPERATIONS = new Set([
  "edit_paragraph",
  "replace_paragraph",
  "modify_text",
  "replace_range",
]);

// Repair: strip stray [P#] markers the model may have leaked into document text.
const PARAGRAPH_MARKER_RE = /\[P\d+\]\s*/g;

// Detect JSON schema field names leaking into content fields.
const SCHEMA_LEAK_RE = /\bparagraphIndex\b|\bendParagraphIndex\b|"operation"\s*:/;

/**
 * Sanitize and validate a raw AI change array. Repairs what can be repaired
 * (e.g. stray [P#] markers) and rejects what cannot, with a machine-readable
 * reason per rejection.
 *
 * @param {Array} rawChanges
 * @param {number} paragraphCount - total paragraphs in the document (0/unknown
 *   disables the upper-bound index check but still rejects index < 1)
 * @param {string[]|null} [paragraphTexts] - optional source paragraphs for exact find validation
 * @returns {{ changes: Array, rejected: Array<{change: *, paragraphIndex: *, operation: *, reason: string}> }}
 */
export function sanitizeChangeSet(rawChanges, paragraphCount, paragraphTexts = null) {
  const changes = [];
  const rejected = [];
  if (!Array.isArray(rawChanges)) {
    return { changes, rejected };
  }

  const haveCount = Number.isInteger(paragraphCount) && paragraphCount > 0;
  const seen = new Set(); // `${paragraphIndex}:${operation}` of kept changes

  for (const raw of rawChanges) {
    const reject = (reason) =>
      rejected.push({
        change: raw,
        paragraphIndex: raw && typeof raw === "object" ? raw.paragraphIndex : undefined,
        operation: raw && typeof raw === "object" ? raw.operation : undefined,
        reason,
      });

    // 1. Drop non-objects and unrecognized operations.
    if (
      !raw ||
      typeof raw !== "object" ||
      Array.isArray(raw) ||
      !VALID_OPERATIONS.has(raw.operation)
    ) {
      reject("invalid_operation");
      continue;
    }

    const operation = raw.operation;
    const idx = raw.paragraphIndex;

    // 2. Index range (clamp is unsafe for edits, so we reject).
    // Content-bearing ops may target one past the last paragraph (paragraphCount+1)
    // to APPEND new content at the end of the document; the apply engine handles it.
    const appendAllowed =
      operation === "replace_paragraph" ||
      operation === "replace_range" ||
      operation === "edit_paragraph";
    const maxIndex = appendAllowed ? paragraphCount + 1 : paragraphCount;
    if (!Number.isInteger(idx) || idx < 1 || (haveCount && idx > maxIndex)) {
      reject("index_out_of_range");
      continue;
    }
    if (operation === "replace_range") {
      const endIdx = raw.endParagraphIndex;
      if (!Number.isInteger(endIdx) || endIdx < idx || (haveCount && endIdx > maxIndex)) {
        reject("index_out_of_range");
        continue;
      }
    }

    // 3. Strip [P#] markers from content fields (repair, on a shallow copy).
    const change = { ...raw };
    if (change.replacements != null) {
      if (operation !== "edit_paragraph" || (typeof change.newContent === "string" && change.newContent.trim() !== "")) {
        reject("invalid_replacements");
        continue;
      }
      const shapeError = validateLocalizedReplacements(change.replacements);
      if (shapeError) {
        reject(shapeError);
        continue;
      }
      change.replacements = change.replacements.map(item => ({
        find: item.find,
        replace: item.replace.replace(PARAGRAPH_MARKER_RE, ""),
        ...(item.occurrence != null ? { occurrence: item.occurrence } : {})
      }));
    }
    for (const field of ["content", "newContent", "replacementText"]) {
      if (typeof change[field] === "string") {
        change[field] = change[field].replace(PARAGRAPH_MARKER_RE, "");
      }
    }

    // 3b. Normalize wrong-field content and strip inapplicable fields (repair).
    // Models (esp. thinking models) sometimes put the payload in the wrong field
    // or fill unused fields with meta-commentary. The apply engine falls back to
    // `content ?? newContent ?? replacementText`, so leaked junk in an unused
    // field could otherwise be written INTO the document. We repair the common
    // wrong-field mix-ups (content<->newContent), then delete every field that
    // does not belong to the chosen operation so junk can never reach the engine.
    const hasText = (v) => typeof v === "string" && v.trim() !== "";
    if (operation === "replace_paragraph" || operation === "replace_range") {
      if (!hasText(change.content) && hasText(change.newContent)) {
        change.content = change.newContent;
      }
      delete change.newContent;
      delete change.originalText;
      delete change.replacementText;
    } else if (operation === "edit_paragraph") {
      if (change.replacements == null && !hasText(change.newContent) && hasText(change.content)) {
        change.newContent = change.content;
      }
      delete change.content;
      delete change.originalText;
      delete change.replacementText;
      if (change.replacements != null) delete change.newContent;
    } else if (operation === "modify_text") {
      delete change.content;
      delete change.newContent;
    }

    // 4. Empty-content guard for replace_paragraph / replace_range / edit_paragraph.
    if (operation === "replace_paragraph" || operation === "replace_range") {
      if (!hasText(change.content)) {
        reject("empty_content");
        continue;
      }
    }
    if (operation === "edit_paragraph" && change.replacements == null && !hasText(change.newContent)) {
      reject("empty_content");
      continue;
    }
    if (change.replacements != null && Array.isArray(paragraphTexts)) {
      const actual = paragraphTexts[idx - 1];
      const anchor = normalizeForAnchor(change.anchorText);
      const anchorMatches = anchor === "" || normalizeForAnchor(actual).includes(anchor);
      if (anchorMatches) {
        const reason = validateLocalizedReplacements(change.replacements, actual);
        if (reason) {
          reject(reason);
          continue;
        }
      }
    }

    // 5/6. modify_text specific guards.
    if (operation === "modify_text") {
      if (typeof change.originalText === "string" && change.originalText.length > 80) {
        reject("original_text_too_long");
        continue;
      }
      const replacement = change.replacementText;
      if (
        typeof replacement === "string" &&
        (/\n/.test(replacement) ||
          /\|.*\|/.test(replacement) ||
          /^\s*([-*]|\d+\.)\s/m.test(replacement))
      ) {
        reject("modify_text_structural_content");
        continue;
      }
    }

    // 7. Pseudo-table guard: a single-line pipe string is plain text, not a table.
    if (operation === "replace_paragraph" || operation === "replace_range") {
      const content = change.content;
      if (
        typeof content === "string" &&
        content.includes("|") &&
        !content.includes("\n") &&
        !/\|\s*-{3,}/.test(content)
      ) {
        reject("malformed_table");
        continue;
      }
    }

    // 8. Schema-text leak guard.
    let leaked = false;
    const contentValues = [change.content, change.newContent, change.replacementText,
      ...(change.replacements || []).map(item => item.replace)];
    for (const value of contentValues) {
      if (typeof value === "string" && SCHEMA_LEAK_RE.test(value)) {
        leaked = true;
        break;
      }
    }
    if (leaked) {
      reject("schema_text_in_content");
      continue;
    }

    // 9. Dedupe on (paragraphIndex, operation); keep the first.
    const key = `${idx}:${operation}`;
    if (seen.has(key)) {
      reject("duplicate_target");
      continue;
    }
    seen.add(key);

    changes.push(change);
  }

  return { changes, rejected };
}

/**
 * Parse a JSON array, salvaging the complete leading objects when the text was
 * truncated mid-stream (e.g. a model repetition loop exhausted maxOutputTokens
 * and the JSON ends mid-string). Returns the parsed array, or null when nothing
 * parseable can be recovered.
 *
 * Only top-level `[ {...}, {...} ]` arrays of objects are supported (the diff
 * change-set shape). String contents (including escaped quotes and braces) are
 * handled correctly by the scanner.
 *
 * @param {string} text
 * @returns {{ changes: Array|null, repaired: boolean }}
 */
export function repairTruncatedJsonArray(text) {
  if (typeof text !== "string" || text.trim() === "") {
    return { changes: null, repaired: false };
  }

  try {
    const parsed = JSON.parse(text);
    return { changes: Array.isArray(parsed) ? parsed : null, repaired: false };
  } catch {
    // Fall through to salvage.
  }

  const start = text.indexOf("[");
  if (start < 0) {
    return { changes: null, repaired: false };
  }

  let depth = 0; // depth 1 = inside the top-level array
  let inString = false;
  let escaped = false;
  let lastCompleteObjectEnd = -1;

  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (ch === "\\") {
        escaped = true;
      } else if (ch === '"') {
        inString = false;
      }
      continue;
    }
    if (ch === '"') {
      inString = true;
    } else if (ch === "[" || ch === "{") {
      depth += 1;
    } else if (ch === "]" || ch === "}") {
      depth -= 1;
      if (ch === "}" && depth === 1) {
        lastCompleteObjectEnd = i;
      }
    }
  }

  if (lastCompleteObjectEnd < 0) {
    return { changes: null, repaired: false };
  }

  try {
    const salvaged = JSON.parse(`${text.slice(start, lastCompleteObjectEnd + 1)}]`);
    return { changes: Array.isArray(salvaged) ? salvaged : null, repaired: true };
  } catch {
    return { changes: null, repaired: false };
  }
}

// Cap the per-change rejection lines shown to the model; a degenerate change set
// (e.g. one change repeated dozens of times) should not blow up the retry prompt.
const MAX_REJECTION_LINES = 8;

// Actionable, model-facing guidance per rejection reason code.
const REJECTION_HINTS = {
  anchor_mismatch:
    "the anchorText did not match the targeted paragraph; copy the first 30-60 characters of the correct paragraph verbatim",
  invalid_operation:
    "operation must be one of edit_paragraph, replace_paragraph, modify_text, replace_range",
  index_out_of_range:
    "paragraphIndex (and endParagraphIndex for replace_range) must be within the document",
  empty_content:
    'replace_paragraph/replace_range require non-empty "content"; edit_paragraph requires non-empty "newContent" or "replacements"',
  invalid_replacements:
    'edit_paragraph replacements require a non-empty array of {find, replace} items, with no newContent; occurrence must be a positive integer when provided',
  replacement_find_not_found:
    'each find must match the current target paragraph exactly, including case and punctuation; an occurrence must exist',
  replacement_ambiguous_find:
    'find appears more than once in the target paragraph; specify its 1-based occurrence',
  original_text_too_long:
    'modify_text "originalText" must be 80 characters or fewer; use edit_paragraph with the full rewritten paragraph instead',
  modify_text_structural_content:
    "modify_text cannot insert line breaks, list markers, or tables; use edit_paragraph or replace_paragraph",
  malformed_table:
    "table content must be a full multiline Markdown table (header row, |---| separator, data rows), not a single pipe-delimited line",
  schema_text_in_content:
    "content fields must contain only document text, not JSON field names like paragraphIndex or operation",
  duplicate_target: "a duplicate change for the same paragraph and operation was dropped",
};

/**
 * Build a model-facing, one-line-per-rejection summary. Handles both anchor
 * rejections (verifyAnchor) and sanitizer rejections (sanitizeChangeSet).
 *
 * @param {Array} rejected
 * @returns {string}
 */
export function formatRejections(rejected) {
  if (!Array.isArray(rejected) || rejected.length === 0) {
    return "";
  }
  const lines = rejected.map((r) => {
    const where = r.paragraphIndex != null ? `P${r.paragraphIndex}` : "unknown paragraph";
    const op = r.operation || "unknown";
    const hint = REJECTION_HINTS[r.reason] ? ` (${REJECTION_HINTS[r.reason]})` : "";
    let detail = "";
    if (r.reason === "anchor_mismatch") {
      const claimed = r.anchorText ? `you claimed it starts with "${r.anchorText}"` : "";
      const actual = r.actualTextSnippet
        ? `; it actually starts with "${r.actualTextSnippet}"`
        : "";
      const joined = [claimed, actual].filter(Boolean).join("");
      detail = joined ? ` — ${joined}` : "";
    }
    return `- ${where} (${op}): ${r.reason}${hint}${detail}.`;
  });

  if (lines.length > MAX_REJECTION_LINES) {
    const shown = lines.slice(0, MAX_REJECTION_LINES);
    shown.push(`- …and ${lines.length - MAX_REJECTION_LINES} more rejection(s) like the above.`);
    return shown.join("\n");
  }
  return lines.join("\n");
}
