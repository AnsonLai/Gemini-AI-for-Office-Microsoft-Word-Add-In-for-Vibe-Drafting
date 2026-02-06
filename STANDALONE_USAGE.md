# Standalone Reconciliation Usage

This project now includes a standalone entrypoint that avoids Word JS APIs:

- `src/taskpane/modules/reconciliation/standalone.js`

## 1) Configure XML Provider

In browsers, native `DOMParser` / `XMLSerializer` are used automatically.

In Node.js, configure an XML provider first:

```js
import { configureXmlProvider } from './src/taskpane/modules/reconciliation/standalone.js';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

configureXmlProvider({ DOMParser, XMLSerializer });
```

## 2) Read `.docx` and extract XML

```js
import JSZip from 'jszip';
import fs from 'node:fs/promises';

const input = await fs.readFile('./input.docx');
const zip = await JSZip.loadAsync(input);
const documentXml = await zip.file('word/document.xml').async('string');
```

## 3) Apply Reconciliation

```js
import { applyRedlineToOxml } from './src/taskpane/modules/reconciliation/standalone.js';

const originalText = 'The quick brown fox.';
const modifiedText = 'The **quick** brown fox jumps.';

const result = await applyRedlineToOxml(documentXml, originalText, modifiedText, {
  author: 'AI Assistant',
  generateRedlines: true
});
```

## 4) Write XML back into `.docx`

```js
zip.file('word/document.xml', result.oxml);
const output = await zip.generateAsync({ type: 'nodebuffer' });
await fs.writeFile('./output.docx', output);
```

## 5) Optional logger configuration

```js
import { configureLogger } from './src/taskpane/modules/reconciliation/standalone.js';

configureLogger({
  log: () => {},
  warn: () => {},
  error: console.error
});
```
