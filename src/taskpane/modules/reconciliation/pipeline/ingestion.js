/**
 * OOXML Reconciliation Pipeline - Ingestion Facade
 *
 * Backward-compatible exports for paragraph and table ingestion.
 */

export { ingestOoxml, ingestParagraphElement, detectNumberingContext } from './ingestion-paragraph.js';
export { ingestTableToVirtualGrid } from './ingestion-table.js';
