
import { applyRedlineToOxml } from './src/taskpane/modules/reconciliation/oxml-engine.js';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import fs from 'fs';

// Polyfill DOMParser/XMLSerializer
global.DOMParser = DOMParser;
global.XMLSerializer = XMLSerializer;

// Polyfill NodeList iterator
const dummyDoc = new DOMParser().parseFromString('<root/>', 'text/xml');
const nodeListProto = Object.getPrototypeOf(dummyDoc.childNodes);
if (!nodeListProto[Symbol.iterator]) {
  nodeListProto[Symbol.iterator] = function* () {
    for (let i = 0; i < this.length; i++) {
      yield this[i];
    }
  };
}

const initialXmlPartial = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:rPr>
                 <w:rStyle w:val="Normal"/>
              </w:rPr>
              <w:t>Hello World</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

const initialXmlOff = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:rPr>
                 <w:b w:val="0"/>
              </w:rPr>
              <w:t>Hello World</w:t>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;

async function runTest() {
  console.log("--- Test 6: Partial Formatting (Split + Order) ---");
  // Should behave as "Hello " + "**World**"
  const partialText = "Hello **World**";
  try {
    const res = await applyRedlineToOxml(initialXmlPartial, "Hello World", partialText);

    // Check 1: Bold tag exists
    if (!res.oxml.includes('<w:b/>') && !res.oxml.includes('<w:b>')) {
      console.log("FAIL: Bold tag missing in partial update");
      writeFail("Test6_Missing", res.oxml);
    } else {
      // Check 2: Order - rStyle should come BEFORE b
      // Regex to find w:rPr content using multiline
      const match = res.oxml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/g);
      let passedOrder = false;
      if (match) {
        for (const rPrContent of match) {
          if (rPrContent.includes('<w:b') && rPrContent.includes('w:rStyle')) {
            const styleIndex = rPrContent.indexOf('w:rStyle');
            const boldIndex = rPrContent.indexOf('<w:b');
            if (boldIndex < styleIndex) {
              console.log("FAIL: <w:b> appeared BEFORE <w:rStyle>. Invalid OOXML.");
              writeFail("Test6_Order", res.oxml);
            } else {
              passedOrder = true;
            }
          }
        }
      }
      if (passedOrder) console.log("PASS: Partial formatting applied with correct order");
      else if (!match) console.log("FAIL: No rPr matched?");
      else console.log("FAIL: Bold tag found but not checked against rStyle (maybe split wrong?)");
    }
  } catch (e) {
    console.error("Test 6 Failed:", e);
  }

  console.log("\n--- Test 7: Overriding 'Off' Property ---");
  const boldText = "**Hello World**";
  try {
    const res = await applyRedlineToOxml(initialXmlOff, "Hello World", boldText);

    // We want to see a bold tag that is NOT val="0" (or val="false")
    const matches = [...res.oxml.matchAll(/<w:b(?: [^>]*)?\/>/g)];
    const tags = matches.map(m => m[0]);

    // Check if any tag acts as "on"
    // <w:b/> -> on
    // <w:b w:val="1"/> -> on
    // <w:b w:val="0"/> -> off

    const hasEnable = tags.some(t => !t.includes('w:val="0"') && !t.includes('w:val="false"'));

    if (hasEnable) {
      console.log("PASS: Bold tag added/updated despite existing disable tag.");
    } else {
      console.log("FAIL: Only found existing disable-bold tag. Formatting was NOT applied.");
      console.log("Tags found:", tags);
      writeFail("Test7_Override", res.oxml);
    }
  } catch (e) {
    console.error("Test 7 Failed:", e);
  }
}

function writeFail(name, content) {
  fs.writeFileSync(`${name}_failed.xml`, content);
  console.log(`Wrote ${name}_failed.xml`);
}

runTest().catch(console.error);
