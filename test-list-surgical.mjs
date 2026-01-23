/**
 * Test script for multi-paragraph reconciliation
 * Run with: node --experimental-modules test-list-surgical.mjs
 */

// Mock DOMParser and XMLSerializer for Node.js environment
import { JSDOM } from 'jsdom';
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;

import { ingestOoxml } from './src/taskpane/modules/reconciliation/ingestion.js';
import { serializeToOoxml } from './src/taskpane/modules/reconciliation/serialization.js';
import { ReconciliationPipeline } from './src/taskpane/modules/reconciliation/pipeline.js';

// Test 1: Multi-paragraph ingestion
console.log('=== Test 1: Multi-Paragraph Ingestion ===');
const testOoxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>Item One</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>Item Two</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr></w:pPr><w:r><w:t>Item Three</w:t></w:r></w:p>
  </w:body>
</w:document>
`;

const { runModel, acceptedText, pPr } = ingestOoxml(testOoxml);

console.log('Run Model Length:', runModel.length);
console.log('Accepted Text:', JSON.stringify(acceptedText));
console.log('Expected: "Item One\\nItem Two\\nItem Three"');
console.log('Has actual newlines:', acceptedText.includes('\n'));

// Check for PARAGRAPH_START tokens
const paragraphStarts = runModel.filter(r => r.kind === 'paragraph_start');
console.log('PARAGRAPH_START tokens:', paragraphStarts.length);
console.log('Expected: 3');

// Test 2: Round-trip serialization
console.log('\n=== Test 2: Round-Trip Serialization ===');
const serialized = serializeToOoxml(runModel, pPr, []);
console.log('Serialized output has multiple <w:p>:', (serialized.match(/<w:p>/g) || []).length);
console.log('Expected: 3');

// Test 3: Surgical diff on existing list
console.log('\n=== Test 3: Surgical Diff ===');
const pipeline = new ReconciliationPipeline({ author: 'Test', generateRedlines: true });

// Modify only the third item
const modifiedText = 'Item One\nItem Two\nItem Three MODIFIED';

try {
  const result = await pipeline.execute(testOoxml, modifiedText);
  console.log('Result isValid:', result.isValid);

  // Check that deletion of "Item One" is NOT present
  const hasUnwantedDeletion = result.ooxml.includes('<w:delText xml:space="preserve">Item One</w:delText>');
  console.log('Has unwanted deletion of Item One:', hasUnwantedDeletion);
  console.log('Expected: false');

  // Check that insertion of "MODIFIED" IS present
  const hasExpectedInsertion = result.ooxml.includes('MODIFIED');
  console.log('Has expected insertion of MODIFIED:', hasExpectedInsertion);
  console.log('Expected: true');

  if (!hasUnwantedDeletion && hasExpectedInsertion) {
    console.log('\n✅ SUCCESS: Surgical diff is working!');
  } else {
    console.log('\n❌ FAILURE: Still doing full replacement');
  }
} catch (e) {
  console.error('Pipeline error:', e);
}

console.log('\n=== All Tests Complete ===');
