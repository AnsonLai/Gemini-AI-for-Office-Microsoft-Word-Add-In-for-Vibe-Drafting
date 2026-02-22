# @gsd/docx-reconciliation (in-repo prep package)

Host-agnostic OOXML reconciliation engine for `.docx` manipulation with track changes support.

- Primary entrypoint: `index.js`
- Compatibility alias: `standalone.js` (deprecated for new imports)
- Add-in-local entrypoint: `word-addin-entry.js` (not for package publish)
- Architecture guide: `ARCHITECTURE.md`

## Quick Start

```js
import { configureXmlProvider, setDefaultAuthor, setPlatform, applyRedlineToOxml } from './index.js';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

configureXmlProvider({ DOMParser, XMLSerializer });
setDefaultAuthor('My Agent');
setPlatform('Node');

const result = await applyRedlineToOxml(
  paragraphOoxml,
  'Original sentence.',
  'Updated sentence.',
  { generateRedlines: true }
);

console.log(result.hasChanges, result.oxml);
```

## API Overview

- Engine APIs
  - `applyRedlineToOxml`
  - `applyRedlineToOxmlWithListFallback`
  - `reconcileMarkdownTableOoxml`
- Pipeline APIs
  - `ReconciliationPipeline`
  - `ingestWordOoxmlToPlainText`
  - `ingestWordOoxmlToMarkdown`
- Service APIs
  - `generateTableOoxml`
  - comment injection helpers
  - numbering helpers in `services/numbering-helpers.js`
  - package/plumbing helpers in `services/standalone-docx-plumbing.js`
- Orchestration APIs
  - `buildReconciliationPlan`
  - list parsing/markdown/fallback helpers in `orchestration/*`

## Configuration

- `configureXmlProvider({ DOMParser, XMLSerializer })`
  - Required in Node.js runtimes; optional in browsers that already provide DOM APIs.
- `configureLogger({ log, warn, error })`
  - Inject host logger for diagnostics.
- `setDefaultAuthor(author)` and `getDefaultAuthor()`
  - Controls fallback track-change author metadata.
- `setPlatform(platform)` and `getPlatform()`
  - Host-provided platform label used by metadata and diagnostics.

## Hosting Guidance

- Node.js
  - Install and provide `@xmldom/xmldom` to `configureXmlProvider`.
  - Use package/plumbing helpers when editing full `.docx` archives.
- Browser
  - Native `DOMParser` and `XMLSerializer` are usually enough.
  - Use the same host-agnostic entrypoint (`index.js`).
- Word add-in
  - Import add-in integration APIs from `word-addin-entry.js`.
  - Keep Office.js usage in add-in modules, not in core package modules.

## Browser Demo Minimal Example

The browser demo uses the core package plus a host orchestration layer:

1. Read `word/document.xml` from a `.docx` zip (for example with JSZip).
2. Run reconciliation APIs from `index.js`.
3. Apply package helpers:
   - `extractReplacementNodesFromOoxml`
   - `ensureNumberingArtifactsInZip`
   - `ensureCommentsArtifactsInZip`
   - `validateDocxPackage`
4. Write updated XML and package artifacts back to the zip.

Reference implementation: `browser-demo/demo.js`.

## Architecture

See `src/taskpane/modules/reconciliation/ARCHITECTURE.md` for module boundaries, flow, and contributor orientation.
