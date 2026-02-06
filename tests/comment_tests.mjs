import './setup-xml-provider.mjs';
/**
 * Comment Engine Tests
 * 
 * Tests for the OOXML comment injection system.
 * Run with: node --experimental-modules tests/comment_tests.mjs
 */

import { JSDOM } from 'jsdom';
import {
    buildCommentElement,
    buildCommentsPartXml,
    buildCommentMarkers,
    injectCommentsIntoOoxml,
    wrapWithCommentsPart,
    injectCommentsIntoPackage,
    resetRevisionIdCounter
} from '../src/taskpane/modules/reconciliation/services/comment-engine.js';

// --- Mock Browser Environment ---
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.document = dom.window.document;

// --- Test Utilities ---
let passCount = 0;
let failCount = 0;

function assert(condition, message) {
    if (condition) {
        console.log(`  ✅ ${message}`);
        passCount++;
    } else {
        console.log(`  ❌ ${message}`);
        failCount++;
    }
}

function runTest(name, testFn) {
    console.log(`\n=== Test: ${name} ===`);
    resetRevisionIdCounter(1000);
    try {
        testFn();
    } catch (e) {
        console.log(`  💥 ERROR: ${e.message}`);
        failCount++;
    }
}

// --- Sample OOXML ---
const simpleParagraphOoxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r><w:t>Hello world, this is a test paragraph.</w:t></w:r>
        </w:p>
    </w:body>
</w:document>
`;

const multiParagraphOoxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r><w:t>First paragraph content.</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Second paragraph has different text.</w:t></w:r>
        </w:p>
        <w:p>
            <w:r><w:t>Third paragraph concludes the document.</w:t></w:r>
        </w:p>
    </w:body>
</w:document>
`;

const multiRunParagraphOoxml = `
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:r><w:t>Hello </w:t></w:r>
            <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
            <w:r><w:t> world!</w:t></w:r>
        </w:p>
    </w:body>
</w:document>
`;

// --- Tests ---

runTest('buildCommentElement creates valid XML', () => {
    const result = buildCommentElement(1, 'Test Author', 'This is a comment', '2025-01-30T00:00:00Z');

    assert(result.includes('w:id="1"'), 'Has comment ID');
    assert(result.includes('w:author="Test Author"'), 'Has author');
    assert(result.includes('This is a comment'), 'Has comment content');
    assert(result.includes('w:initials="TA"'), 'Has initials');
});

runTest('buildCommentElement escapes special characters', () => {
    const result = buildCommentElement(1, 'Test', 'Comment with <special> & "chars"', '2025-01-30T00:00:00Z');

    assert(result.includes('&lt;special&gt;'), 'Escaped angle brackets');
    assert(result.includes('&amp;'), 'Escaped ampersand');
    assert(result.includes('&quot;'), 'Escaped quotes');
});

runTest('buildCommentsPartXml creates complete structure', () => {
    const comments = [
        { id: 1, content: 'Comment 1', author: 'AI', date: '2025-01-30T00:00:00Z' },
        { id: 2, content: 'Comment 2', author: 'AI', date: '2025-01-30T00:00:00Z' }
    ];
    const result = buildCommentsPartXml(comments);

    assert(result.includes('<w:comments'), 'Has comments root element');
    assert(result.includes('w:id="1"'), 'Has first comment');
    assert(result.includes('w:id="2"'), 'Has second comment');
});

runTest('buildCommentMarkers returns all three elements', () => {
    const markers = buildCommentMarkers(42);

    assert(markers.start.includes('commentRangeStart'), 'Has start marker');
    assert(markers.end.includes('commentRangeEnd'), 'Has end marker');
    assert(markers.reference.includes('commentReference'), 'Has reference');
    assert(markers.start.includes('w:id="42"'), 'Start has correct ID');
    assert(markers.end.includes('w:id="42"'), 'End has correct ID');
    assert(markers.reference.includes('w:id="42"'), 'Reference has correct ID');
});

runTest('injectCommentsIntoOoxml - single comment', () => {
    const comments = [
        { paragraphIndex: 1, textToFind: 'test paragraph', commentContent: 'Review this' }
    ];

    const result = injectCommentsIntoOoxml(simpleParagraphOoxml, comments, { author: 'AI' });

    assert(result.commentsApplied === 1, `Applied 1 comment (got ${result.commentsApplied})`);
    assert(result.oxml.includes('commentRangeStart'), 'Has start marker');
    assert(result.oxml.includes('commentRangeEnd'), 'Has end marker');
    assert(result.oxml.includes('commentReference'), 'Has reference');
    assert(result.commentsXml.includes('Review this'), 'Comments XML has content');
});

runTest('injectCommentsIntoOoxml - multiple comments on different paragraphs', () => {
    const comments = [
        { paragraphIndex: 1, textToFind: 'First paragraph', commentContent: 'Comment on first' },
        { paragraphIndex: 2, textToFind: 'Second paragraph', commentContent: 'Comment on second' },
        { paragraphIndex: 3, textToFind: 'Third paragraph', commentContent: 'Comment on third' }
    ];

    const result = injectCommentsIntoOoxml(multiParagraphOoxml, comments, { author: 'AI' });

    assert(result.commentsApplied === 3, `Applied 3 comments (got ${result.commentsApplied})`);
    assert(result.warnings.length === 0, `No warnings (got ${result.warnings.length})`);
});

runTest('injectCommentsIntoOoxml - text not found', () => {
    const comments = [
        { paragraphIndex: 1, textToFind: 'nonexistent text', commentContent: 'This should fail' }
    ];

    const result = injectCommentsIntoOoxml(simpleParagraphOoxml, comments, { author: 'AI' });

    assert(result.commentsApplied === 0, 'No comments applied');
    assert(result.warnings.length === 1, 'Has warning');
    assert(result.warnings[0].includes('Could not find'), 'Warning mentions not found');
});

runTest('injectCommentsIntoOoxml - paragraph out of range', () => {
    const comments = [
        { paragraphIndex: 99, textToFind: 'test', commentContent: 'Out of range' }
    ];

    const result = injectCommentsIntoOoxml(simpleParagraphOoxml, comments, { author: 'AI' });

    assert(result.commentsApplied === 0, 'No comments applied');
    assert(result.warnings.length === 1, 'Has warning');
    assert(result.warnings[0].includes('out of range'), 'Warning mentions out of range');
});

runTest('injectCommentsIntoOoxml - text spanning multiple runs', () => {
    const comments = [
        { paragraphIndex: 1, textToFind: 'bold world', commentContent: 'Spans runs' }
    ];

    const result = injectCommentsIntoOoxml(multiRunParagraphOoxml, comments, { author: 'AI' });

    assert(result.commentsApplied === 1, `Applied 1 comment (got ${result.commentsApplied})`);
    assert(result.oxml.includes('commentRangeStart'), 'Has start marker');
    assert(result.oxml.includes('commentRangeEnd'), 'Has end marker');
});

runTest('injectCommentsIntoOoxml - empty comments array', () => {
    const result = injectCommentsIntoOoxml(simpleParagraphOoxml, [], { author: 'AI' });

    assert(result.commentsApplied === 0, 'No comments applied');
    assert(result.oxml === simpleParagraphOoxml, 'Original OXML unchanged');
});

runTest('wrapWithCommentsPart creates complete package', () => {
    const documentXml = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>';
    const commentsXml = '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="AI"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments>';

    const result = wrapWithCommentsPart(documentXml, commentsXml);

    assert(result.includes('pkg:package'), 'Has package wrapper');
    assert(result.includes('/word/comments.xml'), 'Has comments part');
    assert(result.includes('/word/document.xml'), 'Has document part');
    assert(result.includes('relationships/comments'), 'Has comments relationship');
});

runTest('Unique comment IDs across multiple injections', () => {
    resetRevisionIdCounter(100);

    const comments1 = [{ paragraphIndex: 1, textToFind: 'Hello', commentContent: 'First' }];
    const comments2 = [{ paragraphIndex: 1, textToFind: 'world', commentContent: 'Second' }];

    const result1 = injectCommentsIntoOoxml(simpleParagraphOoxml, comments1);
    const result2 = injectCommentsIntoOoxml(simpleParagraphOoxml, comments2);

    // Extract IDs from results
    const idMatch1 = result1.oxml.match(/commentRangeStart w:id="(\d+)"/);
    const idMatch2 = result2.oxml.match(/commentRangeStart w:id="(\d+)"/);

    assert(idMatch1 && idMatch2, 'Both have IDs');
    assert(idMatch1[1] !== idMatch2[1], `IDs are unique (${idMatch1[1]} vs ${idMatch2[1]})`);
});

// Sample pkg:package structure (simplified from getOoxml output)
const samplePackageOoxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p><w:r><w:t>Test document</w:t></w:r></w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

runTest('injectCommentsIntoPackage adds comments part to existing package', () => {
    const commentsXml = '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="1" w:author="AI"><w:p><w:r><w:t>Test comment</w:t></w:r></w:p></w:comment></w:comments>';

    const result = injectCommentsIntoPackage(samplePackageOoxml, commentsXml);

    assert(result.includes('/word/comments.xml'), 'Has comments part');
    assert(result.includes('Test comment'), 'Comments content is included');
    assert(result.includes('relationships/comments'), 'Has comments relationship');
    assert(result.includes('rId2'), 'New relationship has unique ID');
    assert(result.includes('/word/document.xml'), 'Original document part preserved');
    assert(result.includes('Test document'), 'Original content preserved');
});

// --- Summary ---
console.log('\n========================================');
console.log(`RESULTS: ${passCount} passed, ${failCount} failed`);
console.log('========================================\n');

if (failCount > 0) {
    process.exit(1);
}

