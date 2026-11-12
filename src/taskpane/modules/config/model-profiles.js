/** Host-neutral provider limits. Keep the existing catalogue, caps and preview warnings. */

const COMMON = {
  maxOutputTokens: 48000,
  retries: 3,
  previewThrottleWarning: false,
};

const MODEL_PROFILES = {
  "gemini-2.5-pro": { ...COMMON },
  "gemini-2.5-flash": { ...COMMON },
  "gemini-flash-latest": { ...COMMON },
  "gemini-flash-lite-latest": { ...COMMON },
  // 3.x are preview/throttled relative to 2.5; the timeout warning telling users
  // to revert to 2.5 only makes sense for these models.
  "gemini-3.8-flash": { ...COMMON, previewThrottleWarning: true },
  "gemini-3.1-pro-preview": { ...COMMON, previewThrottleWarning: true },
  "gemini-4-carbon": { ...COMMON },
  "gemini-4-argon": { ...COMMON },
};

// Unknown/unlisted models keep the same provider limits.
const DEFAULT_PROFILE = { ...COMMON };

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
