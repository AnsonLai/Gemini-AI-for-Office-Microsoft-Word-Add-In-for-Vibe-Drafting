/**
 * model-profiles.js
 *
 * One place that declares how each supported Gemini model should be driven, so
 * per-model configuration is centralized instead of scattered as inline
 * model-name conditionals across the codebase.
 *
 * This module is intentionally free of Office.js / DOM / localStorage access so
 * it can be imported in plain Node and unit tested.
 *
 * IMPORTANT: the values below are chosen to PRESERVE current runtime behavior:
 *   - maxOutputTokens 48000 matches the previous API_LIMITS.MAX_OUTPUT_TOKENS.
 *   - temperature 0.1 matches the deterministic temperature used by the diff /
 *     structured-output call (callGeminiForDiffs).
 *   - retries 3 matches the previous callGeminiWithRetry default.
 * The agentic chat loop historically did NOT set a temperature; we keep it that
 * way (it does not consume `temperature` from the profile).
 *
 * Profile fields:
 *   - maxOutputTokens {number}        generationConfig.maxOutputTokens
 *   - temperature {number}            temperature for structured/diff generation
 *   - retries {number}               max attempts for callGeminiWithRetry
 *   - toolCallReliability {"high"|"medium"|"low"}  qualitative signal for future routing
 *   - supportsResponseSchema {boolean}             responseSchema honored by the model
 *   - previewThrottleWarning {boolean}             show the "preview / throttled,
 *                                                  revert to 2.5" message on timeout
 */

const COMMON = {
  maxOutputTokens: 48000,
  // Tighter budget for the diff-generation call: a legitimate change set never
  // needs 48k tokens, and a smaller cap bounds the cost/latency of model
  // repetition loops (observed on gemini-3.8-flash) before salvage kicks in.
  diffMaxOutputTokens: 16384,
  temperature: 0.1,
  retries: 3,
  toolCallReliability: "high",
  supportsResponseSchema: true,
  previewThrottleWarning: false,
};

const MODEL_PROFILES = {
  "gemini-2.5-pro": { ...COMMON },
  "gemini-2.5-flash": { ...COMMON },
  "gemini-flash-latest": { ...COMMON },
  "gemini-flash-lite-latest": { ...COMMON, toolCallReliability: "medium" },
  // 3.x are preview/throttled relative to 2.5; the timeout warning telling users
  // to revert to 2.5 only makes sense for these models.
  "gemini-3.8-flash": { ...COMMON, previewThrottleWarning: true },
  "gemini-3.1-pro-preview": { ...COMMON, previewThrottleWarning: true },
};

// Used for unknown/unlisted models. Slightly more conservative reliability signal.
const DEFAULT_PROFILE = { ...COMMON, toolCallReliability: "medium" };

/**
 * Resolve the profile for a model name. Falls back to a prefix match so
 * versioned names (e.g. "gemini-2.5-flash-002") resolve to their base profile,
 * then to DEFAULT_PROFILE.
 *
 * @param {string} modelName
 * @returns {object} a model profile (never null)
 */
export function getModelProfile(modelName) {
  if (typeof modelName !== "string" || modelName === "") {
    return DEFAULT_PROFILE;
  }
  if (MODEL_PROFILES[modelName]) {
    return MODEL_PROFILES[modelName];
  }
  // Prefer the longest matching prefix so "gemini-3.8-flash-preview" resolves to
  // "gemini-3.8-flash" rather than a shorter accidental match.
  const key = Object.keys(MODEL_PROFILES)
    .filter((k) => modelName.startsWith(k))
    .sort((a, b) => b.length - a.length)[0];
  return key ? MODEL_PROFILES[key] : DEFAULT_PROFILE;
}

export { MODEL_PROFILES, DEFAULT_PROFILE };
