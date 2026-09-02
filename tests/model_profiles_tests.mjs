import assert from 'assert';
import {
  getModelProfile,
  MODEL_PROFILES,
  DEFAULT_PROFILE
} from '../src/taskpane/modules/config/model-profiles.js';

function testExactNameLookup() {
  assert.strictEqual(getModelProfile('gemini-2.5-pro'), MODEL_PROFILES['gemini-2.5-pro']);
  assert.strictEqual(getModelProfile('gemini-3.8-flash'), MODEL_PROFILES['gemini-3.8-flash']);
}

function testPrefixMatch() {
  // Versioned name resolves to the base profile.
  assert.strictEqual(getModelProfile('gemini-2.5-flash-002'), MODEL_PROFILES['gemini-2.5-flash']);
  // Longest prefix wins (3.5-flash-preview -> 3.5-flash, not a shorter accidental match).
  assert.strictEqual(getModelProfile('gemini-3.8-flash-preview'), MODEL_PROFILES['gemini-3.8-flash']);
}

function testUnknownReturnsDefault() {
  assert.strictEqual(getModelProfile('totally-unknown-model'), DEFAULT_PROFILE);
  assert.strictEqual(getModelProfile(''), DEFAULT_PROFILE);
  assert.strictEqual(getModelProfile(undefined), DEFAULT_PROFILE);
  assert.strictEqual(getModelProfile(null), DEFAULT_PROFILE);
}

function testBehaviorPreservingDefaults() {
  // These values must match prior hardcoded behavior so WP3 introduces no change.
  for (const name of ['gemini-2.5-pro', 'gemini-2.5-flash', 'gemini-flash-latest']) {
    const p = getModelProfile(name);
    assert.strictEqual(p.maxOutputTokens, 48000, `${name} maxOutputTokens`);
    assert.strictEqual(p.diffMaxOutputTokens, 16384, `${name} diffMaxOutputTokens`);
    assert.strictEqual(p.temperature, 0.1, `${name} temperature`);
    assert.strictEqual(p.retries, 3, `${name} retries`);
  }
}

function testPreviewThrottleFlag() {
  // Only 3.x preview models carry the "revert to 2.5" throttle warning.
  assert.strictEqual(getModelProfile('gemini-3.8-flash').previewThrottleWarning, true);
  assert.strictEqual(getModelProfile('gemini-3.1-pro-preview').previewThrottleWarning, true);
  assert.strictEqual(getModelProfile('gemini-2.5-pro').previewThrottleWarning, false);
  assert.strictEqual(getModelProfile('gemini-2.5-flash').previewThrottleWarning, false);
}

testExactNameLookup();
testPrefixMatch();
testUnknownReturnsDefault();
testBehaviorPreservingDefaults();
testPreviewThrottleFlag();

console.log('model_profiles_tests passed');
