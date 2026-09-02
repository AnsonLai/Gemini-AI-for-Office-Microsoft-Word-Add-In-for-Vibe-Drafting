# Reconciliation Core: Repository Extraction & Publish Plan

> **Migrated on 2026-08-29:** Remaining publication and repository-boundary
> work was consolidated into [`2026-08-29-package-boundaries-and-integrations.md`](../2026-08-29-package-boundaries-and-integrations.md).
> This document is retained in `migrated/` as historical detail.

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Extract the reconciliation core from `src/taskpane/modules/reconciliation/` into a standalone Git repository, set it up for npm publishing and CDN distribution, and update the current AIWordPlugin repo to consume it as an external dependency. The package must work cleanly via:

1. **npm** (`npm install @ansonlai/docx-redline-js`)
2. **CDN** (`import ... from 'https://esm.sh/@ansonlai/docx-redline-js'`)
3. **Local git clone** (`import ... from './docx-redline-js/index.js'`)

**Guiding Principle:** Ship the raw ES module source as the primary distribution. Add a pre-bundled ESM file (with `diff-match-patch` inlined) for CDN consumers who don't want import maps. No transpilation — the codebase uses only baseline ES2020+ features that all modern runtimes support.

---

## Current State (Post-Extraction Prep)

The previous plan (`2026-02-21-reconciliation-core-extraction.md`) completed all 12 tasks:

- **Entry points normalized**: `index.js` (primary, host-agnostic), `standalone.js` (deprecated shim re-exporting `index.js`), `word-addin-entry.js` (add-in-only, excluded from package)
- **All host/AI contamination removed**: No `Office.*` globals, no hardcoded `'Gemini AI'` — all configurable via `adapters/config.js`
- **Helpers extracted**: `services/numbering-helpers.js`, `engine/formatting-removal.js`, core targeting helpers moved to `core/`
- **Tests split**: `tests/core/` (13 tests, core-only) and `tests/addin/` (4 tests, integration)
- **Prep `package.json`** exists at `src/taskpane/modules/reconciliation/package.json` with `"private": true`
- **Isolation tests pass**: `no_word_api_standalone_check.mjs` (recursive) and `core_dependency_graph_check.mjs`

### Files Included in Core Package

```
reconciliation/
├── adapters/config.js, logger.js, xml-adapter.js
├── core/types.js, ooxml-identifiers.js, paragraph-offset-policy.js,
│        paragraph-targeting.js, list-targeting.js, table-targeting.js, xml-query.js
├── engine/oxml-engine.js, surgical-mode.js, reconstruction-mode.js,
│         reconstruction-mapper.js, reconstruction-writer.js, format-extraction.js,
│         format-application.js, format-paragraph-targeting.js, format-span-application.js,
│         formatting-removal.js, rpr-helpers.js, run-builders.js, table-cell-context.js, table-mode.js
├── pipeline/pipeline.js, ingestion.js, ingestion-paragraph.js, ingestion-table.js,
│          ingestion-xml.js, ingestion-export.js, content-analysis.js, diff-engine.js,
│          list-generation.js, list-markers.js, patching.js, serialization.js, markdown-processor.js
├── services/comment-builders.js, comment-engine.js, comment-locator.js, comment-package.js,
│           numbering-service.js, numbering-helpers.js, package-builder.js,
│           standalone-docx-plumbing.js, standalone-operation-runner.js, table-reconciliation.js
├── orchestration/route-plan.js, list-markdown.js, list-parsing.js,
│                list-structural-fallback.js, redline-operation-converter.js
├── index.js                 ← Primary entry point
├── standalone.js            ← Deprecated shim (re-exports index.js)
├── package.json             ← Prep artifact (private: true)
├── README.md
└── ARCHITECTURE.md
```

### Files EXCLUDED from Core Package (Stay in AIWordPlugin)

```
reconciliation/
├── integration/integration.js, word-ooxml.js, word-operation-runner.js,
│              word-redline-runner.js, word-route-change.js, word-structured-list.js
└── word-addin-entry.js
```

### Single Runtime Dependency

`diff-match-patch` (^1.0.5) — imported as `import { diff_match_patch } from 'diff-match-patch'` in `pipeline/diff-engine.js`. This is the only non-relative import in the entire core.

### Current Browser Consumption Pattern (demo.html)

```html
<script type="importmap">
  { "imports": { "diff-match-patch": "https://esm.sh/diff-match-patch@1.0.5" } }
</script>
<script type="module" src="./demo.js"></script>
```

`demo.js` imports directly from reconciliation source files via relative paths. The importmap resolves the bare `diff-match-patch` specifier for the browser.

---

## Task 1: Create the New Repository
**Status (2026-02-23): ✅ Completed**

### 1.1 Initialize the Repository

Create a new local directory (sibling to AIWordPlugin, or wherever you prefer):

```bash
mkdir docx-redline-js
cd docx-redline-js
git init
```

### 1.2 Copy Core Package Files

Copy **only** the files that belong to the published package from `src/taskpane/modules/reconciliation/`:

```bash
# From the AIWordPlugin root:
# Copy core directories
cp -r src/taskpane/modules/reconciliation/adapters    ../docx-redline-js/adapters
cp -r src/taskpane/modules/reconciliation/core         ../docx-redline-js/core
cp -r src/taskpane/modules/reconciliation/engine       ../docx-redline-js/engine
cp -r src/taskpane/modules/reconciliation/pipeline     ../docx-redline-js/pipeline
cp -r src/taskpane/modules/reconciliation/services     ../docx-redline-js/services
cp -r src/taskpane/modules/reconciliation/orchestration ../docx-redline-js/orchestration

# Copy root files
cp src/taskpane/modules/reconciliation/index.js        ../docx-redline-js/index.js
cp src/taskpane/modules/reconciliation/standalone.js   ../docx-redline-js/standalone.js
cp src/taskpane/modules/reconciliation/ARCHITECTURE.md ../docx-redline-js/ARCHITECTURE.md
```

**Do NOT copy:**
- `integration/` directory
- `word-addin-entry.js`
- `package.json` (we'll write a new one)
- `README.md` (we'll write a new one)

### 1.3 Copy Core Tests

```bash
mkdir -p ../docx-redline-js/tests
cp tests/setup-xml-provider.mjs ../docx-redline-js/tests/
cp tests/core/*.mjs             ../docx-redline-js/tests/
```

### 1.4 Fix All Import Paths in Copied Tests

The tests currently import from `../../src/taskpane/modules/reconciliation/...`. These must be rewritten to import from the new package root:

**Pattern to find and replace in each test file:**
- `../../src/taskpane/modules/reconciliation/index.js` → `../index.js`
- `../../src/taskpane/modules/reconciliation/` → `../` (for deep imports like `services/standalone-operation-runner.js`)

**Also fix `setup-xml-provider.mjs`:**
- `../src/taskpane/modules/reconciliation/adapters/xml-adapter.js` → `../adapters/xml-adapter.js`

**Also copy any test fixture files** if tests reference sample documents (check for `sample_doc/` references in test files and copy those fixtures).

### 1.5 Verify the Copy Is Clean

```bash
cd ../docx-redline-js

# Should find NO references to Office/Word API
grep -r "Office\." --include="*.js" --exclude-dir=node_modules | grep -v "//.*Office"
grep -r "Word\." --include="*.js" --exclude-dir=node_modules | grep -v "//.*Word"

# Should find NO imports escaping the package
grep -rn "from '\.\./\.\." --include="*.js" --exclude-dir=tests
# (tests are allowed to import from ../ since they're one level up from source)

# The ONLY bare specifier import should be diff-match-patch
grep -rn "from '" --include="*.js" --exclude-dir=node_modules | grep -v "from '\." | grep -v "diff-match-patch"
```

---

## Task 2: Write `package.json`
**Status (2026-02-23): ✅ Completed**

Create `docx-redline-js/package.json`:

```json
{
  "name": "@ansonlai/docx-redline-js",
  "version": "0.1.0",
  "description": "Host-independent OOXML reconciliation engine for .docx manipulation with track changes",
  "license": "MIT",
  "type": "module",
  "main": "./index.js",
  "module": "./index.js",
  "exports": {
    ".": {
      "import": "./index.js",
      "default": "./index.js"
    },
    "./standalone": "./standalone.js",
    "./adapters/*": "./adapters/*",
    "./core/*": "./core/*",
    "./engine/*": "./engine/*",
    "./pipeline/*": "./pipeline/*",
    "./services/*": "./services/*",
    "./orchestration/*": "./orchestration/*"
  },
  "files": [
    "adapters/",
    "core/",
    "engine/",
    "pipeline/",
    "services/",
    "orchestration/",
    "index.js",
    "standalone.js",
    "dist/",
    "ARCHITECTURE.md",
    "AGENTS.md",
    "README.md",
    "LICENSE"
  ],
  "dependencies": {
    "diff-match-patch": "^1.0.5"
  },
  "peerDependencies": {
    "@xmldom/xmldom": ">=0.8.0"
  },
  "peerDependenciesMeta": {
    "@xmldom/xmldom": {
      "optional": true
    }
  },
  "devDependencies": {
    "esbuild": "^0.24.0",
    "@xmldom/xmldom": "^0.9.0"
  },
  "scripts": {
    "build": "node scripts/build.mjs",
    "test": "node scripts/run-tests.mjs",
    "test:isolation": "node tests/no_word_api_standalone_check.mjs && node tests/core_dependency_graph_check.mjs",
    "prepublishOnly": "npm run test:isolation && npm run build"
  },
  "keywords": [
    "docx",
    "ooxml",
    "reconciliation",
    "track-changes",
    "redlines",
    "word",
    "office",
    "document",
    "xml"
  ],
  "repository": {
    "type": "git",
    "url": "https://github.com/YOUR_ORG/docx-redline-js.git"
  },
  "engines": {
    "node": ">=18.0.0"
  }
}
```

**Key design decisions:**
- `"type": "module"` — package is ESM-only (no CommonJS build). Node 18+ is the baseline.
- `exports` map gives granular access to subdirectories for deep imports like `@ansonlai/docx-redline-js/services/standalone-operation-runner.js`.
- `dist/` directory will contain the CDN-ready pre-bundled file (built in Task 3).
- `@xmldom/xmldom` is an optional peer dep — browsers don't need it, Node.js does.
- `devDependencies` include `esbuild` for building the CDN bundle and `@xmldom/xmldom` for running tests locally.

---

## Task 3: Add CDN-Ready Bundle Build
**Status (2026-02-23): ✅ Completed**

### 3.1 Create Build Script

Create `docx-redline-js/scripts/build.mjs`:

```js
import { build } from 'esbuild';
import { readFileSync } from 'fs';

const pkg = JSON.parse(readFileSync('./package.json', 'utf8'));

// ESM bundle with diff-match-patch inlined (for CDN/browser <script type="module">)
await build({
  entryPoints: ['./index.js'],
  bundle: true,
  format: 'esm',
  outfile: 'dist/docx-redline-js.esm.js',
  platform: 'neutral',         // no Node builtins assumed
  target: 'es2020',
  minify: false,                // keep readable for debugging
  sourcemap: true,
  banner: {
    js: `// @ansonlai/docx-redline-js v${pkg.version} — https://github.com/YOUR_ORG/docx-redline-js`
  },
  external: ['@xmldom/xmldom']  // never bundle the Node-only XML parser
});

// Minified version for production CDN use
await build({
  entryPoints: ['./index.js'],
  bundle: true,
  format: 'esm',
  outfile: 'dist/docx-redline-js.esm.min.js',
  platform: 'neutral',
  target: 'es2020',
  minify: true,
  sourcemap: true,
  external: ['@xmldom/xmldom']
});

console.log('Build complete: dist/docx-redline-js.esm.js, dist/docx-redline-js.esm.min.js');
```

**What this produces:**
- `dist/docx-redline-js.esm.js` — Single-file ESM bundle with `diff-match-patch` inlined. Can be loaded from a CDN with a bare `<script type="module">` or `import()`. Readable for debugging.
- `dist/docx-redline-js.esm.min.js` — Minified version for production.
- Both have source maps.
- `@xmldom/xmldom` is kept external since CDN consumers use the browser's native `DOMParser`.

### 3.2 Create `.gitignore`

Create `docx-redline-js/.gitignore`:

```
node_modules/
dist/
*.tgz
.DS_Store
```

`dist/` is gitignored but included in `"files"` — npm pack will build it via `prepublishOnly`.

### 3.3 Verify Build

```bash
cd docx-redline-js
npm install
npm run build
ls -la dist/
# Should see: docx-redline-js.esm.js, docx-redline-js.esm.min.js, and source maps
```

Verification run on 2026-02-23 (local): `npm install` and `npm run build` both succeeded, and `dist/` contains:
- `docx-redline-js.esm.js`
- `docx-redline-js.esm.js.map`
- `docx-redline-js.esm.min.js`
- `docx-redline-js.esm.min.js.map`

---

## Task 4: Add Test Runner Script
**Status (2026-02-23): ✅ Completed**

Create `docx-redline-js/scripts/run-tests.mjs`:

```js
import { readdirSync } from 'fs';
import { join } from 'path';
import { execSync } from 'child_process';

const testDir = join(import.meta.dirname, '..', 'tests');
const testFiles = readdirSync(testDir)
  .filter(f => f.endsWith('.mjs') && f !== 'setup-xml-provider.mjs')
  .sort();

let passed = 0;
let failed = 0;
const failures = [];

for (const file of testFiles) {
  const filePath = join(testDir, file);
  process.stdout.write(`  ${file} ... `);
  try {
    execSync(`node "${filePath}"`, { stdio: 'pipe', timeout: 30000 });
    console.log('PASS');
    passed++;
  } catch (err) {
    console.log('FAIL');
    failures.push({ file, stderr: err.stderr?.toString() || err.message });
    failed++;
  }
}

console.log(`\n${passed} passed, ${failed} failed out of ${passed + failed} tests`);
if (failures.length > 0) {
  for (const f of failures) {
    console.error(`\n--- ${f.file} ---\n${f.stderr}`);
  }
  process.exit(1);
}
```

**Verification:**

```bash
npm test
# All 13+ core tests should pass
```

Verification run on 2026-02-23 (local): `npm test` succeeded with `13 passed, 0 failed out of 13 tests`.

---

## Task 5: Write `README.md`
**Status (2026-02-23): ✅ Completed**

Create `docx-redline-js/README.md` — the primary documentation for all consumption patterns. This must be comprehensive enough that someone can get started without reading source code.

Structure:

```markdown
# @ansonlai/docx-redline-js

Host-independent OOXML reconciliation engine for `.docx` manipulation with track changes (redlines).

Converts AI-generated or programmatic text/markdown edits into valid Office Open XML (OOXML) with proper `w:ins`/`w:del` revision markup that Word displays as native tracked changes.

## Features

- **Text reconciliation** with word-level diffing and native-looking redlines
- **Formatting** (bold, italic, underline, strikethrough) via surgical `w:rPrChange`
- **Lists** — generate and edit real Word lists (`w:numPr`) from markdown
- **Tables** — virtual-grid diffing for cell-level edits with merge safety
- **Comments** — inject OOXML comments anchored to specific text ranges
- **Highlights** — apply highlight colors to runs
- **Markdown ↔ OOXML** — bidirectional: ingest OOXML to markdown, convert markdown to OOXML
- **Package plumbing** — helpers for numbering.xml, comments.xml, content types, and relationship wiring
- **Zero host dependencies** — works in Node.js, browsers, Deno, and any JS runtime with a DOM parser

## Install

### npm / Node.js
```bash
npm install @ansonlai/docx-redline-js
```

### CDN (browser `<script type="module">`)
```html
<script type="module">
  import { applyRedlineToOxml } from 'https://esm.sh/@ansonlai/docx-redline-js';
</script>
```

Or use the pre-bundled file (no import map needed, `diff-match-patch` is inlined):
```html
<script type="module">
  import { applyRedlineToOxml } from 'https://cdn.jsdelivr.net/npm/@ansonlai/docx-redline-js/dist/docx-redline-js.esm.min.js';
</script>
```

### Local git clone
```bash
git clone https://github.com/YOUR_ORG/docx-redline-js.git
```
```js
import { applyRedlineToOxml } from './docx-redline-js/index.js';
```

## Quick Start

### Node.js
```js
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
  configureXmlProvider, setDefaultAuthor,
  applyRedlineToOxml
} from '@ansonlai/docx-redline-js';

// One-time setup
configureXmlProvider({ DOMParser, XMLSerializer });
setDefaultAuthor('My App');

// Apply a text edit with tracked changes
const result = await applyRedlineToOxml(
  paragraphOoxml,           // Original OOXML of the paragraph
  'Original sentence.',     // Current visible text
  'Updated sentence.',      // Desired new text
  { generateRedlines: true, author: 'Editor' }
);

console.log(result.hasChanges);  // true
console.log(result.oxml);        // OOXML with w:ins/w:del markup
```

### Browser
```js
import {
  setDefaultAuthor,
  applyRedlineToOxml
} from '@ansonlai/docx-redline-js';
// Browser has native DOMParser — no configureXmlProvider needed

setDefaultAuthor('Browser Editor');

const result = await applyRedlineToOxml(oxml, original, modified, {
  generateRedlines: true
});
```

## API Reference

### Configuration (call once at startup)

| Function | Purpose |
|----------|---------|
| `configureXmlProvider({ DOMParser, XMLSerializer })` | Inject XML parser. Required in Node.js; browsers have native support. |
| `configureLogger({ log, warn, error })` | Replace default console logger. |
| `setDefaultAuthor(name)` | Set fallback track-change author (default: `'Author'`). |
| `setPlatform(label)` | Set platform label for diagnostics (default: `'Unknown'`). |

### Engine (primary reconciliation APIs)

| Function | Purpose |
|----------|---------|
| `applyRedlineToOxml(oxml, original, modified, options)` | Core engine: reconcile text/markdown edit into OOXML with optional redlines. |
| `applyRedlineToOxmlWithListFallback(oxml, original, modified, options)` | Same as above, with automatic single-line list structural fallback. |
| `reconcileMarkdownTableOoxml(oxml, original, markdownTable, options)` | Table-specific reconciliation convenience wrapper. |

### Pipeline (lower-level access)

| Function | Purpose |
|----------|---------|
| `ReconciliationPipeline` | Class for direct pipeline access (ingestion, diff, patch, serialize). |
| `ingestWordOoxmlToPlainText(oxml)` | Extract readable text from OOXML (tags stripped). |
| `ingestWordOoxmlToMarkdown(oxml)` | Convert OOXML to markdown (headings, bold/italic, lists). |
| `ingestOoxml(oxml)` | Flatten OOXML into a `RunModel` with offset mapping. |
| `preprocessMarkdown(text)` | Clean markdown and extract format hints. |

### Services

| Function | Purpose |
|----------|---------|
| `injectCommentsIntoOoxml(oxml, comments, options)` | Add comments anchored to text ranges. |
| `generateTableOoxml(headers, rows, options)` | Generate a new `w:tbl` from data. |
| `createDynamicNumberingIdState(numberingXml)` | Collision-safe numbering ID allocator. |
| `ensureNumberingArtifactsInZip(zip, numberingXml)` | Merge numbering into a .docx package. |
| `ensureCommentsArtifactsInZip(zip, commentsXml)` | Merge comments into a .docx package. |
| `validateDocxPackage(zip)` | Validate .docx structural consistency. |

### Deep Imports

For advanced use cases, import directly from submodules:
```js
import { applyOperationToDocumentXml } from '@ansonlai/docx-redline-js/services/standalone-operation-runner.js';
import { getParagraphText } from '@ansonlai/docx-redline-js/core/paragraph-targeting.js';
```

## Working with .docx Files

This package operates on OOXML strings (the XML inside `.docx` archives), not on `.docx` files directly. To work with full `.docx` files:

1. **Extract** the `.docx` zip (using [JSZip](https://stuk.github.io/jszip/), [fflate](https://github.com/101arrowz/fflate), or Node's `zlib`)
2. **Read** `word/document.xml` from the archive
3. **Apply** reconciliation APIs to the XML string
4. **Use** package plumbing helpers to merge numbering/comments back into the archive
5. **Write** the modified archive back to a `.docx` file

```js
import JSZip from 'jszip';
import {
  configureXmlProvider, applyRedlineToOxml,
  parseXmlStrictStandalone, getBodyElementFromDocument,
  ensureNumberingArtifactsInZip, validateDocxPackage
} from '@ansonlai/docx-redline-js';

const zip = await JSZip.loadAsync(docxBuffer);
const documentXml = await zip.file('word/document.xml').async('string');

// ... apply edits with applyRedlineToOxml ...
// ... merge artifacts with ensureNumberingArtifactsInZip(zip, numberingXml) ...

const output = await zip.generateAsync({ type: 'blob' });
```

## Architecture

See [ARCHITECTURE.md](./ARCHITECTURE.md) for module layout, data flow, and contributor guidance.

See [AGENTS.md](./AGENTS.md) for a condensed structural overview designed for AI agents building on top of this package.
```

---

## Task 6: Write `AGENTS.md`
**Status (2026-02-23): ✅ Completed**

Create `docx-redline-js/AGENTS.md` — a concise structural overview designed specifically for AI coding agents (Claude, Copilot, Cursor, etc.) to quickly understand the package without reading the entire codebase.

```markdown
# AGENTS.md — AI Agent Quick Reference

> This file helps AI coding agents understand @ansonlai/docx-redline-js quickly.
> Read this instead of exploring the full source tree.

## What This Package Does

Converts text/markdown edits into valid Office Open XML (OOXML) with Word-native tracked changes. Feed it original OOXML + desired text → get back OOXML with `w:ins`/`w:del` revision markup.

## Conceptual Model

```
Input: (paragraph OOXML, original text, modified text, options)
  ↓
Engine routes to: format-only | surgical | reconstruction | list | table mode
  ↓
Output: { oxml: string, hasChanges: boolean, warnings?: string[] }
```

The engine works at the **paragraph level**. For full-document operations, callers iterate paragraphs and call the engine per-paragraph (or use the operation runner for batch operations).

## Entry Point

```js
import { applyRedlineToOxml, configureXmlProvider } from '@ansonlai/docx-redline-js';
```

`index.js` is the single entry point. All public APIs are exported from here.

## Required Setup (Node.js only)

```js
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
configureXmlProvider({ DOMParser, XMLSerializer });
```

Browsers have native DOM APIs — no setup needed.

## Key APIs by Use Case

### "I want to apply a text edit with tracked changes"
```js
const result = await applyRedlineToOxml(oxml, originalText, modifiedText, {
  generateRedlines: true,
  author: 'Agent Name'
});
// result.oxml contains the modified OOXML
// result.hasChanges tells you if anything changed
```

### "I want to apply a text edit without tracked changes"
```js
const result = await applyRedlineToOxml(oxml, originalText, modifiedText, {
  generateRedlines: false
});
```

### "I want to convert OOXML to readable text or markdown"
```js
import { ingestWordOoxmlToPlainText, ingestWordOoxmlToMarkdown } from '@ansonlai/docx-redline-js';
const plainText = ingestWordOoxmlToPlainText(documentXml);
const markdown = ingestWordOoxmlToMarkdown(documentXml);
```

### "I want to add a comment to a paragraph"
```js
import { injectCommentsIntoOoxml } from '@ansonlai/docx-redline-js';
const result = injectCommentsIntoOoxml(paragraphOoxml, [
  { text: 'Review this clause', targetText: 'force majeure', author: 'Agent' }
]);
```

### "I want to apply multiple operations to a full document.xml"
```js
import { applyOperationToDocumentXml } from '@ansonlai/docx-redline-js/services/standalone-operation-runner.js';
const result = await applyOperationToDocumentXml(documentXml, operation, options);
// operation = { type: 'redline', targetRef: 'P3', modified: 'New text', ... }
```

### "I want to convert a text paragraph into a real Word list"
```js
const result = await applyRedlineToOxml(oxml, 'Item text', '1. Item text', {
  generateRedlines: true
});
// Engine auto-detects list markers and generates w:numPr structure
```

### "I want to edit a table"
```js
import { reconcileMarkdownTableOoxml } from '@ansonlai/docx-redline-js';
const result = await reconcileMarkdownTableOoxml(tableOoxml, originalText, markdownTable);
```

### "I need to work with a full .docx file"
This package operates on XML strings, not .docx files. Use JSZip or similar to:
1. Extract `word/document.xml` from the .docx zip
2. Apply changes with this package
3. Use `ensureNumberingArtifactsInZip()` and `ensureCommentsArtifactsInZip()` to merge generated artifacts
4. Write the zip back to .docx

## Module Map

```
index.js                    ← All exports. Start here.
adapters/
  config.js                 ← setDefaultAuthor(), setPlatform()
  xml-adapter.js            ← configureXmlProvider() for DOM parser injection
  logger.js                 ← configureLogger() for custom logging
core/
  types.js                  ← Enums (DiffOp, RunKind), constants (NS_W), revision helpers
  paragraph-targeting.js    ← Find/match paragraphs by reference or text
  list-targeting.js         ← List scope detection and insertion planning
  table-targeting.js        ← Table scope heuristics
engine/
  oxml-engine.js            ← Main router: applyRedlineToOxml()
  surgical-mode.js          ← In-place edits for structure-sensitive content
  reconstruction-mode.js    ← Full paragraph rebuild from diff
  format-application.js     ← Bold/italic/underline via w:rPrChange
  formatting-removal.js     ← Remove formatting from runs
  table-mode.js             ← Table reconciliation/text-to-table
pipeline/
  pipeline.js               ← ReconciliationPipeline class (5-stage)
  ingestion.js              ← OOXML → RunModel
  ingestion-export.js       ← OOXML → plain text / markdown
  diff-engine.js            ← Word-level diffing (uses diff-match-patch)
  markdown-processor.js     ← Strip markers, extract format hints
  serialization.js          ← RunModel → OOXML, pkg:package wrapping
  list-generation.js        ← Generate list paragraphs from markdown
services/
  standalone-operation-runner.js  ← Batch operation runner for full documents
  standalone-docx-plumbing.js     ← .docx package wiring helpers
  numbering-helpers.js            ← Numbering ID allocation and merging
  comment-engine.js               ← Comment injection
  table-reconciliation.js         ← Virtual grid table diffing
  package-builder.js              ← pkg:package fragment builders
orchestration/
  route-plan.js             ← Classify content into apply strategies
  list-markdown.js          ← Build list markdown from structured data
  list-structural-fallback.js ← Force text→list conversion when diff is no-op
```

## Common Patterns

### Options object
Most APIs accept an `options` object:
```js
{
  generateRedlines: true,   // false = direct edit, true = tracked changes
  author: 'Name',           // Track-change author (falls back to configured default)
}
```

### Return shape
Engine APIs return:
```js
{
  oxml: string,             // Modified OOXML
  hasChanges: boolean,      // Whether any modifications were made
  warnings?: string[],      // Non-fatal issues
  numberingXml?: string,    // Generated numbering definitions (for list operations)
  useNativeApi?: boolean    // true = standalone can't handle this (format-only edge case)
}
```

### OOXML wrapping
OOXML fragments must be wrapped in `pkg:package` for Word's `insertOoxml` API:
```js
import { wrapInDocumentFragment } from '@ansonlai/docx-redline-js';
const wrapped = wrapInDocumentFragment(rawOoxml, { includeNumbering: true, numberingXml });
```

## Gotchas

1. **Always call `configureXmlProvider` first in Node.js** — otherwise all parse/serialize calls will fail silently.
2. **`applyRedlineToOxml` is async** — it awaits internal pipeline stages.
3. **The engine works on paragraph OOXML**, not full document XML. For document-level operations, use `standalone-operation-runner.js`.
4. **List operations may return `numberingXml`** — you must merge this into the document's `word/numbering.xml` for lists to render.
5. **`useNativeApi: true` in results** means the operation requires Word's native API (format-only edge case). Standalone callers get a no-op with a warning instead.
```

---

## Task 7: Write `ARCHITECTURE.md` (Package Version)
**Status (2026-02-23): ✅ Completed**

Update the existing `ARCHITECTURE.md` (copied from the prep phase) to be self-contained for the new repo. Remove references to the AIWordPlugin parent project, `word-addin-entry.js`, and `integration/`.

Key changes from the current version:
- Remove "Add-in local only" scope note
- Remove `word-addin-entry.js` from entry points
- Rewrite "End-to-End Flow" to only describe the host-agnostic path
- Add a section on the build output (`dist/`)
- Add a section on testing (`npm test`)
- Keep the "Fast Orientation For Contributors" section — it's valuable

---

## Task 8: Add `LICENSE` File
**Status (2026-02-23): ✅ Completed**

Create `docx-redline-js/LICENSE` with the MIT license text (matching `package.json` license field). Use the current year (2026) and whatever copyright holder name is appropriate.

---

## Task 9: Create Initial Commit and Tag
**Status (2026-02-23): ✅ Completed**

```bash
cd docx-redline-js
npm install
npm run build
npm test

git add .
git commit -m "Initial release: @ansonlai/docx-redline-js v0.1.0

Host-independent OOXML reconciliation engine extracted from AIWordPlugin.
Supports text, formatting, list, table, and comment reconciliation
with Word-native tracked changes."

git tag v0.1.0
```

Verification run on 2026-02-23 (local):
- `npm install` completed (dependencies already up to date)
- `npm run build` succeeded
- `npm test` succeeded (`13 passed, 0 failed`)
- Initial commit created: `70ca9b6`
- Tag created: `v0.1.0`

---

## Task 10: Update AIWordPlugin to Consume the New Package
**Status (2026-02-23): ✅ Completed**
**Status Update (2026-02-23): ✅ Scope renamed from `@gsd` to `@ansonlai` across AIWordPlugin and the extracted package repo.**
**Status Update (2026-02-23): ✅ Naming convention normalized from `docx-reconciliation` to `docx-redline-js` (package name, import specifiers, integration path names, MCP service name, docs/examples).**
**Verification Update (2026-02-23):**
- `Docx Redline JS`: `npm install`, `npm run build`, and `npm test` passed with package name `@ansonlai/docx-redline-js`.
- `AIWordPlugin`: `npm install`, `npm run build`, and add-in integration tests passed after import/dependency scope updates.
- `mcp/docx-server`: `npm install --prefix mcp/docx-server` completed with `@ansonlai/docx-redline-js` dependency.

### 10.1 Install the Package Locally

While the package is not yet published to npm, use a local file reference:

```bash
# In AIWordPlugin root
npm install ../docx-redline-js
```

This adds `"@ansonlai/docx-redline-js": "file:../docx-redline-js"` to `package.json`. Later, when published, change to `"@ansonlai/docx-redline-js": "^0.1.0"`.

### 10.2 Move `integration/` to AIWordPlugin-Local Module

The `integration/` directory currently lives inside the reconciliation folder. It needs to stay in AIWordPlugin but import from the package instead of relative paths.

**Steps:**
1. Move `src/taskpane/modules/reconciliation/integration/` to `src/taskpane/modules/docx-redline-js-integration/` (a new sibling directory, outside the package)
2. Move `src/taskpane/modules/reconciliation/word-addin-entry.js` to `src/taskpane/modules/docx-redline-js-integration/index.js`
3. In every file in `docx-redline-js-integration/`:
   - Replace relative imports like `'../engine/oxml-engine.js'` with package imports like `'@ansonlai/docx-redline-js/engine/oxml-engine.js'`
   - Replace imports from `'../core/...'` with `'@ansonlai/docx-redline-js/core/...'`
   - etc. for all `../` imports that pointed into the reconciliation package
4. The `docx-redline-js-integration/index.js` (formerly `word-addin-entry.js`) should re-export from `@ansonlai/docx-redline-js` plus the local integration modules

### 10.3 Update Consumer Import Paths

**`agentic-tools.js`** (and any other add-in command modules):
- Change: `from '../reconciliation/word-addin-entry.js'`
- To: `from '../docx-redline-js-integration/index.js'`

**`browser-demo/demo.js`**:
- Change: `from '../src/taskpane/modules/reconciliation/index.js'`
- To: `from '@ansonlai/docx-redline-js'` (or keep relative path to local clone, depending on preference)
- Change: `from '../src/taskpane/modules/reconciliation/services/standalone-operation-runner.js'`
- To: `from '@ansonlai/docx-redline-js/services/standalone-operation-runner.js'`

**Note for browser-demo:** If using `@ansonlai/docx-redline-js` package imports in the browser, you need an import map:
```html
<script type="importmap">
{
  "imports": {
    "@ansonlai/docx-redline-js": "../node_modules/@ansonlai/docx-redline-js/index.js",
    "@ansonlai/docx-redline-js/": "../node_modules/@ansonlai/docx-redline-js/",
    "diff-match-patch": "https://esm.sh/diff-match-patch@1.0.5"
  }
}
</script>
```

Alternatively, the browser-demo can use the CDN bundle directly:
```html
<script type="module">
import { applyRedlineToOxml } from 'https://cdn.jsdelivr.net/npm/@ansonlai/docx-redline-js/dist/docx-redline-js.esm.min.js';
</script>
```

**MCP server** (`mcp/docx-server/src/services/docx-redline-js-service.mjs`):
- Change: `from '../../../../src/taskpane/modules/reconciliation/index.js'`
- To: `from '@ansonlai/docx-redline-js'`
- Also add `"@ansonlai/docx-redline-js": "file:../../../../docx-redline-js"` to `mcp/docx-server/package.json` dependencies (or install from npm once published)

### 10.4 Remove the In-Repo Reconciliation Source

After all consumers are updated and tests pass:

1. Delete `src/taskpane/modules/reconciliation/` (all of it — the package lives in its own repo now)
2. Keep `src/taskpane/modules/docx-redline-js-integration/` (the Word-specific adapter layer)
3. Move `tests/core/` test files to the new package repo (they already were copied in Task 1)
4. Keep `tests/addin/` (these test the integration layer)
5. Update `tests/addin/` imports to use `@ansonlai/docx-redline-js` instead of relative paths to the now-deleted reconciliation directory

### 10.5 Update AIWordPlugin Documentation

- Update `ARCHITECTURE.md` to reference `@ansonlai/docx-redline-js` as an external package
- Update `STATE.md` to note the extraction is complete
- Update `ROADMAP.md` to mark repository split as done

### 10.6 Verify Everything Still Works

```bash
# AIWordPlugin
npm run build                                   # webpack should resolve @ansonlai/docx-redline-js
node tests/addin/integration_tests.mjs          # add-in integration
node tests/addin/word_operation_runner_adapter_tests.mjs
node tests/addin/shared_operation_bridge_tests.mjs
node tests/addin/migrated_tool_cutover_tests.mjs

# docx-redline-js (in separate repo)
npm test                                        # all core tests

# Manual: browser-demo/demo.html
# Manual: npm run mcp:docx
```

Verification run on 2026-02-23 (local):
- `npm run build` succeeded (webpack warnings only: bundle size recommendations)
- `node tests/addin/integration_tests.mjs` passed
- `node tests/addin/word_operation_runner_adapter_tests.mjs` passed
- `node tests/addin/shared_operation_bridge_tests.mjs` passed
- `node tests/addin/migrated_tool_cutover_tests.mjs` passed

---

## Task 11: Publish to npm (When Ready)
**Status (2026-02-23): ⏳ In Progress — blocked on npm authentication (`ENEEDAUTH`).**
**Status Update (2026-02-23):**
- `npm whoami` failed: access token expired/revoked; machine is not authenticated to npm.
- `npm view @ansonlai/docx-redline-js version` returned `404` (package not yet published).
- `npm pack --dry-run` succeeded after cleanup; tarball now includes only `dist/docx-redline-js.*` assets (legacy `dist/docx-reconciliation.*` artifacts removed).
- `npm publish --access public` was attempted; `prepublishOnly` checks passed (`test:isolation` + `build`), then publish failed only due to npm auth.
- Follow-up publish attempt after re-login reached registry publish step but failed with `E403`:
  - "Two-factor authentication or granular access token with bypass 2fa enabled is required to publish packages."
- Retry publish attempt (2026-02-23) produced the same `E403` requirement; no email-based OTP flow was triggered by npm CLI in this environment.
- GitHub release-gated CI/CD workflows added in package repo (`.github/workflows/ci.yml`, `.github/workflows/publish.yml`) with publish restricted to `release.published` events and tag/version validation.

### 11.1 Create an npm Organization

If `@ansonlai` scope is not yet registered:
```bash
npm login
npm org create ansonlai
```

### 11.2 First Publish

```bash
cd docx-redline-js
# Verify package contents
npm pack --dry-run

# Publish (prepublishOnly runs tests + build automatically)
npm publish --access public
```

### 11.3 Update AIWordPlugin to Use Published Package

```bash
cd AIWordPlugin
npm install @ansonlai/docx-redline-js@0.1.0
# Remove the file: reference from package.json
```

### 11.4 CDN Availability

After publishing to npm, the package is automatically available on:

- **esm.sh**: `https://esm.sh/@ansonlai/docx-redline-js` (auto-bundles ESM for browser)
- **jsDelivr**: `https://cdn.jsdelivr.net/npm/@ansonlai/docx-redline-js/dist/docx-redline-js.esm.min.js`
- **unpkg**: `https://unpkg.com/@ansonlai/docx-redline-js/dist/docx-redline-js.esm.min.js`
- **Skypack**: `https://cdn.skypack.dev/@ansonlai/docx-redline-js`

**Browser consumers can use either:**

Option A — Import map with esm.sh (auto-resolves `diff-match-patch`):
```html
<script type="importmap">
  { "imports": { "@ansonlai/docx-redline-js": "https://esm.sh/@ansonlai/docx-redline-js" } }
</script>
<script type="module">
  import { applyRedlineToOxml } from '@ansonlai/docx-redline-js';
</script>
```

Option B — Pre-bundled (no import map, `diff-match-patch` inlined):
```html
<script type="module">
  import { applyRedlineToOxml } from 'https://cdn.jsdelivr.net/npm/@ansonlai/docx-redline-js/dist/docx-redline-js.esm.min.js';
</script>
```

Option C — Download locally:
```bash
npm pack @ansonlai/docx-redline-js
tar -xzf ansonlai-docx-redline-js-0.1.0.tgz
# Use package/ directory
```

---

## Execution Order

```
Task 1  (create repo + copy files)          ← first
Task 2  (write package.json)                ← depends on Task 1
Task 3  (add CDN bundle build)              ← depends on Task 2
Task 4  (add test runner)                   ← depends on Task 1
Task 5  (write README.md)                   ← depends on Task 2
Task 6  (write AGENTS.md)                   ← no deps, can parallel with 5
Task 7  (update ARCHITECTURE.md)            ← no deps, can parallel with 5-6
Task 8  (add LICENSE)                       ← no deps, can parallel
Task 9  (initial commit + tag)              ← depends on Tasks 1-8
Task 10 (update AIWordPlugin consumers)     ← depends on Task 9
Task 11 (publish to npm)                    ← depends on Task 10 passing
```

**Suggested parallel batches:**
1. Tasks 1-4 (repo setup, package.json, build, tests)
2. Tasks 5-8 in parallel (documentation + license)
3. Task 9 (commit)
4. Task 10 (update consumers)
5. Task 11 (publish)

---

## Verification Checklist

### New Package Repo (`docx-redline-js/`)
```bash
npm install                    # installs diff-match-patch + esbuild + @xmldom/xmldom
npm run build                  # produces dist/docx-redline-js.esm.js + .min.js
npm test                       # all 13+ core tests pass
npm run test:isolation         # boundary checks pass
npm pack --dry-run             # verify included files look correct
```

### AIWordPlugin (After Task 10)
```bash
npm run build                  # webpack resolves @ansonlai/docx-redline-js
node tests/addin/*.mjs         # all add-in tests pass
# Manual: browser-demo works
# Manual: MCP server works
```

### CDN/Browser Smoke Test (After Task 11)
Create a minimal HTML file to verify CDN consumption:
```html
<!DOCTYPE html>
<html>
<body>
<script type="module">
import { applyRedlineToOxml, setDefaultAuthor } from
  'https://cdn.jsdelivr.net/npm/@ansonlai/docx-redline-js/dist/docx-redline-js.esm.min.js';

setDefaultAuthor('CDN Test');
const oxml = `<?xml version="1.0"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
<pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>Hello world</w:t></w:r></w:p></w:body>
</w:document></pkg:xmlData></pkg:part></pkg:package>`;

const result = await applyRedlineToOxml(oxml, 'Hello world', 'Hello updated world', {
  generateRedlines: true
});
document.body.textContent = `hasChanges: ${result.hasChanges}, oxml length: ${result.oxml.length}`;
</script>
</body>
</html>
```

---

## Notes for the Executing Agent

1. **The reconciliation source is pure ES modules with no transpilation needed.** Do not add Babel, TypeScript compilation, or CommonJS shims. The `"type": "module"` in `package.json` is intentional.

2. **The `standalone.js` shim is kept for backward compatibility.** It's a 5-line file that re-exports from `index.js`. Include it in the published package but document it as deprecated.

3. **`integration/` and `word-addin-entry.js` must NOT be copied to the new repo.** They contain Word API references and belong with the add-in.

4. **`diff-match-patch` is imported as a bare specifier** (`from 'diff-match-patch'`). In the browser without a bundler, this requires either an import map or the pre-bundled `dist/` file. The build script in Task 3 inlines it.

5. **The isolation tests (`no_word_api_standalone_check.mjs` and `core_dependency_graph_check.mjs`) must continue to pass.** They recursively scan all `.js` files for Word API markers and out-of-boundary imports. Run them as part of CI.

6. **When updating AIWordPlugin imports (Task 10),** the webpack config does NOT need changes — webpack resolves `@ansonlai/docx-redline-js` from `node_modules/` automatically. The browser-demo needs an import map update since it doesn't use a bundler.

7. **The browser-demo's `demo.html` already uses an import map for `diff-match-patch`.** When switching to the package, add the package to the same import map. The `diff-match-patch` entry can be removed if using the pre-bundled `dist/` file instead.

8. **Test fixture files:** Some tests may reference sample OOXML documents. Check for file path references in test files (look for `readFileSync`, `sample_doc`, fixture patterns) and copy those files to the new repo's `tests/` directory.
