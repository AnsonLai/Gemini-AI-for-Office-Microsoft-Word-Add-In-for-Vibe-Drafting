/**
 * Standalone reconciliation entrypoint (no Word JS API dependencies).
 */

// Adapters
export { configureXmlProvider } from './adapters/xml-adapter.js';
export { configureLogger } from './adapters/logger.js';

// Engine
export { applyRedlineToOxml, sanitizeAiResponse, parseOoxml, serializeOoxml } from './engine/oxml-engine.js';

// Pipeline components
export { ReconciliationPipeline } from './pipeline/pipeline.js';
export { ingestOoxml } from './pipeline/ingestion.js';
export { preprocessMarkdown } from './pipeline/markdown-processor.js';
export { serializeToOoxml, wrapInDocumentFragment } from './pipeline/serialization.js';

// Comment engine
export {
    injectCommentsIntoOoxml,
    injectCommentsIntoPackage,
    buildCommentElement,
    buildCommentsPartXml
} from './services/comment-engine.js';

// Formatting removal utilities (outside reconciliation folder)
export {
    removeFormattingFromRPr,
    applyFormattingRemovalToOoxml,
    applyHighlightToOoxml
} from '../../ooxml-formatting-removal.js';

// Table/list tools
export { generateTableOoxml } from './services/table-reconciliation.js';
export { NumberingService } from './services/numbering-service.js';

// Core types/constants
export { DiffOp, RunKind, ContainerKind, ContentType, NS_W, escapeXml } from './core/types.js';
