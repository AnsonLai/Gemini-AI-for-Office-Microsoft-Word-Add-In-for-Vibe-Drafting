
import { diff_match_patch } from 'diff-match-patch';
// Adjust path to point to the actual file location
import { wordsToChars, charsToWords } from './src/taskpane/modules/reconciliation/diff-engine.js';

const text1 = "British Columbia";
const text2 = "the State of California";

console.log('--- Direct DMP (Character Level) ---');
console.log(`Original: "${text1}"`);
console.log(`New:      "${text2}"`);

const dmp = new diff_match_patch();
const diffs = dmp.diff_main(text1, text2);
dmp.diff_cleanupSemantic(diffs);
console.log('Result:', JSON.stringify(diffs));

console.log('\n--- Word Level Tokenization ---');
const { chars1, chars2, wordArray } = wordsToChars(text1, text2);
const charDiffs = dmp.diff_main(chars1, chars2);
dmp.diff_cleanupSemantic(charDiffs);
const wordDiffs = charsToWords(charDiffs, wordArray);
console.log('Result:', JSON.stringify(wordDiffs));
