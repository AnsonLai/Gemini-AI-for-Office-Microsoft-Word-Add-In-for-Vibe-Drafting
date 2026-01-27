
import { JSDOM } from 'jsdom';
const dom = new JSDOM('');
global.DOMParser = dom.window.DOMParser;
global.XMLSerializer = dom.window.XMLSerializer;
global.document = dom.window.document;

// Mock the function from oxml-engine.js since it's not exported
function processRunForReconstruction(r, originalFullText) {
    let fullText = originalFullText;
    Array.from(r.childNodes).forEach(rc => {
        if (rc.nodeName === 'w:t') {
            const textContent = rc.textContent || '';
            fullText += textContent;
        }
        // ... checked other types ...
    });
    return fullText;
}

function testExtraction() {
    const parser = new DOMParser();
    // Create a run with br and tab
    const xml = `
    <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:r>
            <w:t>Hello</w:t>
            <w:br/>
            <w:t>World</w:t>
            <w:tab/>
            <w:t>!</w:t>
        </w:r>
    </w:p>
    `;
    const doc = parser.parseFromString(xml, 'text/xml');
    const run = doc.getElementsByTagName('w:r')[0];

    const extracted = processRunForReconstruction(run, '');
    console.log('Extracted Text:', JSON.stringify(extracted));

    if (extracted === 'HelloWorld!') {
        console.log('FAIL: Missing br and tab');
    } else {
        console.log('PASS: Preserved br and tab?');
    }
}

testExtraction();
