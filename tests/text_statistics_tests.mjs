import assert from 'assert';
import {
  stripFormatting,
  countWords,
  countSentences,
  calculateTextStatistics
} from '../src/taskpane/modules/utils/text-statistics.js';

function testStripFormatting() {
  // Test 1: Markdown bold, italic, underline, strikethrough
  const mdFormatted = "This is **bold text**, *italic text*, ++underlined text++, and ~~strikethrough~~.";
  assert.strictEqual(
    stripFormatting(mdFormatted),
    "This is bold text, italic text, underlined text, and strikethrough."
  );

  // Test 2: Internal [P#] anchors
  const internalMarkers = "[P1|Normal] Introduction to the agreement. [P2|ListNumber|L:0|§] Section 1.";
  assert.strictEqual(
    stripFormatting(internalMarkers),
    "Introduction to the agreement. Section 1."
  );

  // Test 3: Markdown headings and blockquotes
  const mdBlocks = "# Main Title\n\n> This is a quote\n\n## Subheading";
  assert.strictEqual(
    stripFormatting(mdBlocks),
    "Main Title\nThis is a quote\nSubheading"
  );

  // Test 4: Markdown links and images
  const mdLinks = "Please check the [Google Website](https://google.com) and ![logo](img.png).";
  assert.strictEqual(
    stripFormatting(mdLinks),
    "Please check the Google Website and logo."
  );

  // Test 5: Markdown lists
  const mdList = "* Item 1\n* Item 2\n1. Numbered item\n2. Second numbered item";
  assert.strictEqual(
    stripFormatting(mdList),
    "Item 1\nItem 2\nNumbered item\nSecond numbered item"
  );

  // Test 6: Markdown tables
  const mdTable = "| Column 1 | Column 2 |\n|---|---|\n| Value A | Value B |";
  assert.strictEqual(
    stripFormatting(mdTable),
    "Column 1 Column 2\nValue A Value B"
  );

  // Test 7: HTML/XML tags and entities
  const htmlText = "Hello &nbsp; <b>world</b> &amp; everyone &lt;tag&gt;!";
  assert.strictEqual(
    stripFormatting(htmlText),
    "Hello world & everyone <tag>!"
  );

  // Test 8: Unicode zero-width chars and non-breaking spaces
  const unicodeText = "Non\u00A0breaking\u200B space\uFEFF test";
  assert.strictEqual(
    stripFormatting(unicodeText),
    "Non breaking space test"
  );
}

function testCountWords() {
  assert.strictEqual(countWords(""), 0);
  assert.strictEqual(countWords("   "), 0);
  assert.strictEqual(countWords("Hello"), 1);
  assert.strictEqual(countWords("Hello world!"), 2);
  assert.strictEqual(countWords("State-of-the-art AI solution"), 6); // State, of, the, art, AI, solution
  assert.strictEqual(countWords("Don't count words wrongly"), 4); // Don't, count, words, wrongly
  assert.strictEqual(countWords("One   two   three\nfour\tfive"), 5);
}

function testCalculateTextStatistics() {
  const input = "[P1|Normal] **The quick brown fox** jumps over the `lazy dog`. ++It was a sunny day.++\n\n[P2|Normal] Second paragraph here.";
  const stats = calculateTextStatistics(input);

  assert.strictEqual(stats.paragraphCount, 2);
  assert.strictEqual(stats.wordCount, 17);
  assert.strictEqual(stats.characterCountWithoutSpaces > 0, true);
  assert.strictEqual(stats.characterCountWithSpaces >= stats.characterCountWithoutSpaces, true);
  assert.strictEqual(typeof stats.estimatedReadingTime, "string");
  assert.strictEqual(typeof stats.preview, "string");
  assert.strictEqual(stats.preview.includes("[P1|Normal]"), false);
  assert.strictEqual(stats.preview.includes("**"), false);
}

function testEmptyInput() {
  const stats = calculateTextStatistics("");
  assert.strictEqual(stats.wordCount, 0);
  assert.strictEqual(stats.characterCountWithSpaces, 0);
  assert.strictEqual(stats.characterCountWithoutSpaces, 0);
  assert.strictEqual(stats.paragraphCount, 0);
  assert.strictEqual(stats.sentenceCount, 0);
}

async function testExecuteGetSelectionStats() {
  const { executeGetSelectionStats } = await import('../src/taskpane/modules/commands/agentic-tools.js');

  // 1. Direct text passed
  const res1 = await executeGetSelectionStats({ text: "**Bold text** with 5 words." });
  assert.strictEqual(res1.success, true);
  assert.strictEqual(res1.stats.wordCount, 5);
  assert.strictEqual(res1.stats.characterCountWithoutSpaces, 19); // "Bold text with 5 words." without spaces

  // 2. Paragraph range from fullDocumentText
  const docContext = "[P1|Normal] First paragraph with **formatted** words.\n[P2|Normal] Second paragraph content.\n[P3|Normal] Third paragraph.";
  const res2 = await executeGetSelectionStats({ startParagraphIndex: 1, endParagraphIndex: 2 }, docContext);
  assert.strictEqual(res2.success, true);
  assert.strictEqual(res2.stats.paragraphCount, 2);
  assert.strictEqual(res2.stats.wordCount, 8);

  // 3. User Highlighted Text from prompt context
  const docWithSelection = "User Highlighted Text:\n\"\"\"Highlighted clause for review.\"\"\"\n\n[P1|Normal] Document text";
  const res3 = await executeGetSelectionStats({}, docWithSelection);
  assert.strictEqual(res3.success, true);
  assert.strictEqual(res3.stats.wordCount, 4);

  // 4. Empty input
  const res4 = await executeGetSelectionStats({}, "");
  assert.strictEqual(res4.success, false);
}

testStripFormatting();
testCountWords();
testCalculateTextStatistics();
testEmptyInput();
await testExecuteGetSelectionStats();

console.log("text_statistics_tests passed successfully!");
