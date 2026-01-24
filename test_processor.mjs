import { preprocessMarkdown } from './src/taskpane/modules/reconciliation/markdown-processor.js';

const testCase1 = "~~NON-DISCLOSURE AGREEMENT~~";
const testCase2 = "++NON-DISCLOSURE AGREEMENT++";
const testCase3 = "**bold** ~~strike~~ ++under++";

console.log("--- Test Case 1: Strikethrough ---");
const res1 = preprocessMarkdown(testCase1);
console.log("Input:", testCase1);
console.log("Clean:", res1.cleanText);
console.log("Hints:", JSON.stringify(res1.formatHints));
if (res1.cleanText === "NON-DISCLOSURE AGREEMENT") {
    console.log("PASS: Strikethrough cleaned");
} else {
    console.log("FAIL: Strikethrough NOT cleaned");
}

console.log("\n--- Test Case 2: Underline ---");
const res2 = preprocessMarkdown(testCase2);
console.log("Input:", testCase2);
console.log("Clean:", res2.cleanText);
console.log("Hints:", JSON.stringify(res2.formatHints));
if (res2.cleanText === "NON-DISCLOSURE AGREEMENT") {
    console.log("PASS: Underline cleaned");
} else {
    console.log("FAIL: Underline NOT cleaned");
}

console.log("\n--- Test Case 3: Mixed ---");
const res3 = preprocessMarkdown(testCase3);
console.log("Input:", testCase3);
console.log("Clean:", res3.cleanText);
console.log("Hints:", JSON.stringify(res3.formatHints));

console.log("\n--- Test Case 4: Nested (Underline + Italic) ---");
const testCase4 = "++*NON-DISCLOSURE AGREEMENT*++";
const res4 = preprocessMarkdown(testCase4);
console.log("Input:", testCase4);
console.log("Clean:", res4.cleanText);
console.log("Hints:", JSON.stringify(res4.formatHints));
if (res4.cleanText === "NON-DISCLOSURE AGREEMENT" && res4.formatHints.length === 2) {
    console.log("PASS: Nested formatting cleaned and hints captured");
} else {
    console.log("FAIL: Nested formatting issue persists");
}
