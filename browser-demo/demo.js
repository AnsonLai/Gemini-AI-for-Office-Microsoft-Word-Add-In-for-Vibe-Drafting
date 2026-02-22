import JSZip from 'https://esm.sh/jszip@3.10.1';
import {
    configureLogger,
    getParagraphText as getParagraphTextFromOxml,
    buildTargetReferenceSnapshot,
    findParagraphByBestTextMatch,
    parseParagraphReference as parseParagraphReferenceShared,
    stripLeadingParagraphMarker as stripLeadingParagraphMarkerShared,
    splitLeadingParagraphMarker as splitLeadingParagraphMarkerShared,
    createDynamicNumberingIdState,
    mergeNumberingXmlBySchemaOrder,
    parseXmlStrictStandalone,
    getBodyElementFromDocument,
    insertBodyElementBeforeSectPr,
    normalizeBodySectionOrderStandalone,
    sanitizeNestedParagraphsInTables,
    ensureNumberingArtifactsInZip,
    ensureCommentsArtifactsInZip,
    validateDocxPackage
} from '../src/taskpane/modules/reconciliation/standalone.js';
import { applyOperationToDocumentXml } from '../src/taskpane/modules/reconciliation/services/standalone-operation-runner.js';

const DEMO_VERSION = '2026-02-15-chat-docx-preview-23';
const GEMINI_API_KEY_STORAGE_KEY = 'browserDemo.geminiApiKey';
const EDIT_MODE_STORAGE_KEY = 'browserDemo.editMode';
const LIBRARY_COLLAPSED_STORAGE_KEY = 'browserDemo.libraryCollapsed';
const LIBRARY_MAX_DOC_CHARS = 12000;
const LIBRARY_MAX_PROMPT_CHARS = 48000;
const GEMINI_REQUEST_PREVIEW_CHARS = 2400;
const DEMO_MARKERS = [
    'DEMO_TEXT_TARGET',
    'DEMO FORMAT TARGET',
    'DEMO_LIST_TARGET',
    'DEMO_TABLE_TARGET'
];
const EDIT_MODE = {
    REDLINE: 'redline',
    DIRECT: 'direct'
};
const ALLOWED_HIGHLIGHT_COLORS = ['yellow', 'green', 'cyan', 'magenta', 'blue', 'red'];
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// docx-preview browser build expects a global JSZip symbol.
if (typeof window !== 'undefined' && !window.JSZip) {
    window.JSZip = JSZip;
}

// ── DOM refs ───────────────────────────────────────────
const fileInput = document.getElementById('docxFile');
const authorInput = document.getElementById('authorInput');
const geminiApiKeyInput = document.getElementById('geminiApiKeyInput');
const saveGeminiKeyBtn = document.getElementById('saveGeminiKeyBtn');
const runBtn = document.getElementById('runBtn');
const logEl = document.getElementById('log');
const logToggle = document.getElementById('logToggle');
const chatMessages = document.getElementById('chat-messages');
const chatInput = document.getElementById('chat-input');
const sendBtn = document.getElementById('sendBtn');
const downloadBtn = document.getElementById('downloadBtn');
const downloadXmlBtn = document.getElementById('downloadXmlBtn');
const docxPreviewEl = document.getElementById('docxPreview');
const previewStatusEl = document.getElementById('previewStatus');
const refreshPreviewBtn = document.getElementById('refreshPreviewBtn');
const editModeInputs = Array.from(document.querySelectorAll('input[name="editMode"]'));
const libraryAddBtn = document.getElementById('libraryAddBtn');
const libraryClearBtn = document.getElementById('libraryClearBtn');
const libraryDocxFilesInput = document.getElementById('libraryDocxFiles');
const libraryDropZone = document.getElementById('libraryDropZone');
const libraryItemsEl = document.getElementById('libraryItems');
const librarySummaryEl = document.getElementById('librarySummary');
const libraryColumnEl = document.querySelector('.library-column');
const libraryToggleBtn = document.getElementById('libraryToggleBtn');

// ── State ──────────────────────────────────────────────
let currentZip = null;           // JSZip instance of the working document
let documentParagraphs = [];     // [{ index, text }] extracted from current docx
let chatHistory = [];            // Gemini multi-turn history [{ role, parts }]
let operationCount = 0;          // total operations applied across turns
let previewRenderer = null;      // docxjs renderAsync function
let previewRenderToken = 0;      // guards stale async preview writes
let editMode = normalizeEditMode(getStoredEditMode());
let libraryDocuments = [];       // [{ id, name, size, paragraphCount, text, originalTextLength, truncated }]
let nextLibraryDocId = 1;
let isLibraryCollapsed = getStoredLibraryCollapsed();
let lastGeminiRequestDebug = null;

function normalizeEditMode(value) {
    return value === EDIT_MODE.DIRECT ? EDIT_MODE.DIRECT : EDIT_MODE.REDLINE;
}

function getEditModeLabel(mode = editMode) {
    return mode === EDIT_MODE.DIRECT ? 'Direct edits' : 'Redlines';
}

function shouldGenerateRedlines(mode = editMode) {
    return normalizeEditMode(mode) !== EDIT_MODE.DIRECT;
}

function formatByteSize(bytes) {
    const size = Number(bytes) || 0;
    if (size < 1024) return `${size} B`;
    const kb = size / 1024;
    if (kb < 1024) return `${kb.toFixed(1)} KB`;
    const mb = kb / 1024;
    return `${mb.toFixed(2)} MB`;
}

function isLikelyDocxFile(file) {
    if (!file) return false;
    const name = String(file.name || '');
    const type = String(file.type || '').toLowerCase();
    return /\.docx$/i.test(name) || type.includes('officedocument.wordprocessingml.document');
}

function truncatePlainText(text, maxChars) {
    const content = String(text || '');
    if (!Number.isInteger(maxChars) || maxChars <= 0 || content.length <= maxChars) {
        return { text: content, truncated: false };
    }
    return {
        text: `${content.slice(0, maxChars).trimEnd()}\n...[truncated]`,
        truncated: true
    };
}

// ── Utility ────────────────────────────────────────────
function log(message) {
    logEl.textContent += `${message}\n`;
    logEl.scrollTop = logEl.scrollHeight;
}

function addMsg(role, html) {
    const el = document.createElement('div');
    el.className = `msg ${role}`;
    el.innerHTML = html;
    chatMessages.appendChild(el);
    chatMessages.scrollTop = chatMessages.scrollHeight;
    return el;
}

function setPreviewStatus(message, level = 'info') {
    if (!previewStatusEl) return;
    previewStatusEl.textContent = message;
    previewStatusEl.classList.remove('success', 'error');
    if (level === 'success') previewStatusEl.classList.add('success');
    if (level === 'error') previewStatusEl.classList.add('error');
}

function getPreviewTimestamp() {
    return new Date().toLocaleTimeString([], {
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    });
}

function logPreviewDiagnostics(stage, error = null) {
    const details = {
        stage,
        hasWindowDocx: typeof window !== 'undefined' && !!window.docx,
        hasRenderAsyncGlobal: typeof window !== 'undefined' && typeof window.docx?.renderAsync === 'function',
        hasWindowJSZip: typeof window !== 'undefined' && !!window.JSZip,
        jszipLoadAsyncType: typeof window !== 'undefined' ? typeof window.JSZip?.loadAsync : 'n/a',
        moduleJsZipLoadAsyncType: typeof JSZip?.loadAsync
    };
    log(`[Preview][Diag] ${JSON.stringify(details)}`);
    console.info('[Preview][Diag]', details);
    if (error) {
        console.error('[Preview][Error]', error);
        if (error?.stack) log(`[Preview][Stack] ${error.stack}`);
    }
}

async function resolvePreviewRenderer() {
    logPreviewDiagnostics('resolvePreviewRenderer:start');
    if (previewRenderer) return previewRenderer;

    if (typeof window.docx?.renderAsync === 'function') {
        previewRenderer = window.docx.renderAsync.bind(window.docx);
        log('[Preview] Using global window.docx.renderAsync');
        return previewRenderer;
    }

    try {
        const module = await import('https://esm.sh/docx-preview@0.3.6');
        const renderAsync = module?.renderAsync || module?.default?.renderAsync;
        if (typeof renderAsync === 'function') {
            previewRenderer = renderAsync;
            log('[Preview] Using dynamically imported docx-preview renderAsync');
            return previewRenderer;
        }
    } catch (error) {
        log(`[WARN] Failed to dynamically load docxjs: ${error?.message || String(error)}`);
        logPreviewDiagnostics('resolvePreviewRenderer:dynamic-import-failed', error);
    }

    logPreviewDiagnostics('resolvePreviewRenderer:unavailable');
    throw new Error('docxjs renderer unavailable');
}

async function renderPreviewFromZip(zip, sourceLabel = 'Document') {
    if (!docxPreviewEl) return;
    const renderToken = ++previewRenderToken;
    setPreviewStatus(`Rendering preview (${sourceLabel})...`);
    if (refreshPreviewBtn) refreshPreviewBtn.disabled = true;
    logPreviewDiagnostics(`renderPreviewFromZip:start:${sourceLabel}`);

    try {
        if (!zip) throw new Error('No document loaded');
        const renderAsync = await resolvePreviewRenderer();
        const blob = await zip.generateAsync({ type: 'blob' });
        const buffer = await blob.arrayBuffer();

        if (renderToken !== previewRenderToken) return;

        docxPreviewEl.replaceChildren();
        await renderAsync(buffer, docxPreviewEl, null, {
            inWrapper: true,
            renderChanges: true,
            renderHeaders: true,
            renderFooters: true,
            renderFootnotes: true,
            renderEndnotes: true,
            useBase64URL: true
        });

        if (renderToken !== previewRenderToken) return;
        setPreviewStatus(`Preview updated (${sourceLabel}) at ${getPreviewTimestamp()}`, 'success');
    } catch (error) {
        const message = error?.message || String(error);
        if (renderToken === previewRenderToken) {
            setPreviewStatus(`Preview failed: ${message}`, 'error');
        }
        log(`[WARN] Preview render failed: ${message}`);
        logPreviewDiagnostics(`renderPreviewFromZip:failed:${sourceLabel}`, error);
    } finally {
        if (renderToken === previewRenderToken && refreshPreviewBtn) {
            refreshPreviewBtn.disabled = !currentZip;
        }
    }
}

// ── Gemini API Key Persistence ─────────────────────────
function getStoredGeminiApiKey() {
    try { return localStorage.getItem(GEMINI_API_KEY_STORAGE_KEY) || ''; }
    catch { return ''; }
}
function setStoredGeminiApiKey(apiKey) {
    try {
        if (apiKey) localStorage.setItem(GEMINI_API_KEY_STORAGE_KEY, apiKey);
        else localStorage.removeItem(GEMINI_API_KEY_STORAGE_KEY);
        return true;
    } catch { return false; }
}

function getStoredEditMode() {
    try { return normalizeEditMode(localStorage.getItem(EDIT_MODE_STORAGE_KEY) || EDIT_MODE.REDLINE); }
    catch { return EDIT_MODE.REDLINE; }
}

function setStoredEditMode(mode) {
    try {
        localStorage.setItem(EDIT_MODE_STORAGE_KEY, normalizeEditMode(mode));
        return true;
    } catch { return false; }
}

function getStoredLibraryCollapsed() {
    try { return localStorage.getItem(LIBRARY_COLLAPSED_STORAGE_KEY) === '1'; }
    catch { return false; }
}

function setStoredLibraryCollapsed(collapsed) {
    try {
        localStorage.setItem(LIBRARY_COLLAPSED_STORAGE_KEY, collapsed ? '1' : '0');
        return true;
    } catch { return false; }
}

function syncEditModeInputs() {
    const normalized = normalizeEditMode(editMode);
    for (const input of editModeInputs) {
        if (!input) continue;
        input.checked = String(input.value || '') === normalized;
    }
}

function setEditMode(nextMode, { announce = true, resetChatHistory = true } = {}) {
    const normalized = normalizeEditMode(nextMode);
    const changed = normalized !== editMode;
    editMode = normalized;
    syncEditModeInputs();
    setStoredEditMode(normalized);
    if (!changed) return;

    if (resetChatHistory) {
        chatHistory = [];
    }

    const modeLabel = getEditModeLabel(normalized);
    log(`[Mode] ${modeLabel} (${shouldGenerateRedlines(normalized) ? 'tracked changes' : 'direct edits'})`);
    if (announce) {
        const suffix = resetChatHistory ? ' Chat context reset.' : '';
        addMsg('system', `Edit mode switched to <strong>${modeLabel}</strong>.${suffix}`);
    }
}

function setLibraryCollapsed(collapsed, { persist = true } = {}) {
    isLibraryCollapsed = !!collapsed;
    if (persist) setStoredLibraryCollapsed(isLibraryCollapsed);

    if (libraryColumnEl) {
        libraryColumnEl.classList.toggle('collapsed', isLibraryCollapsed);
    }
    if (libraryToggleBtn) {
        libraryToggleBtn.textContent = isLibraryCollapsed ? '»' : '«';
        const label = isLibraryCollapsed ? 'Expand library panel' : 'Collapse library panel';
        libraryToggleBtn.title = label;
        libraryToggleBtn.setAttribute('aria-label', label);
        libraryToggleBtn.setAttribute('aria-expanded', String(!isLibraryCollapsed));
    }
}

function renderLibraryList() {
    if (!libraryItemsEl) return;
    libraryItemsEl.replaceChildren();

    if (!Array.isArray(libraryDocuments) || libraryDocuments.length === 0) {
        const emptyEl = document.createElement('div');
        emptyEl.className = 'library-empty';
        emptyEl.textContent = 'Library documents will appear here.';
        libraryItemsEl.appendChild(emptyEl);
        return;
    }

    for (const doc of libraryDocuments) {
        const item = document.createElement('div');
        item.className = 'library-item';

        const top = document.createElement('div');
        top.className = 'library-item-top';

        const head = document.createElement('div');
        head.className = 'library-item-head';

        const includeInput = document.createElement('input');
        includeInput.type = 'checkbox';
        includeInput.className = 'library-item-checkbox';
        includeInput.checked = doc.selected !== false;
        includeInput.dataset.libraryDocToggleId = String(doc.id);
        includeInput.title = 'Include this source in chat prompts';

        const name = document.createElement('div');
        name.className = 'library-item-name';
        name.textContent = doc.name;
        name.title = doc.name;

        const removeBtn = document.createElement('button');
        removeBtn.className = 'library-remove-btn';
        removeBtn.type = 'button';
        removeBtn.textContent = 'Remove';
        removeBtn.dataset.libraryDocId = String(doc.id);

        head.appendChild(includeInput);
        head.appendChild(name);
        top.appendChild(head);
        top.appendChild(removeBtn);

        const meta = document.createElement('div');
        meta.className = 'library-item-meta';
        const truncationLabel = doc.truncated ? ' (truncated)' : '';
        meta.textContent = `${doc.paragraphCount} paragraphs • ${formatByteSize(doc.size)} • ${doc.text.length} chars${truncationLabel}`;

        item.appendChild(top);
        item.appendChild(meta);
        libraryItemsEl.appendChild(item);
    }
}

function renderLibrarySummary() {
    if (!librarySummaryEl) return;
    if (!Array.isArray(libraryDocuments) || libraryDocuments.length === 0) {
        librarySummaryEl.textContent = 'No library docs loaded.';
        return;
    }
    const selectedDocs = getSelectedLibraryDocuments();
    const selectedChars = selectedDocs.reduce((sum, doc) => sum + (doc.text?.length || 0), 0);
    librarySummaryEl.textContent = `${selectedDocs.length}/${libraryDocuments.length} selected • ${selectedChars} context chars`;
}

function refreshLibraryPanel() {
    renderLibrarySummary();
    renderLibraryList();
}

function getSelectedLibraryDocuments() {
    return libraryDocuments.filter(doc => doc.selected !== false);
}

function resetChatHistoryForLibraryChange(actionLabel) {
    chatHistory = [];
    log(`[Library] ${actionLabel}. Chat context reset.`);
}

function setLibraryDocumentSelected(docId, selected, { announce = false } = {}) {
    const targetId = Number(docId);
    const doc = libraryDocuments.find(entry => entry.id === targetId);
    if (!doc) return false;
    const normalized = !!selected;
    if ((doc.selected !== false) === normalized) return false;

    doc.selected = normalized;
    refreshLibraryPanel();
    resetChatHistoryForLibraryChange('Library source selection changed');
    if (announce) {
        const label = normalized ? 'included' : 'excluded';
        addMsg('system', `Reference source <strong>${escapeHtml(doc.name)}</strong> ${label}. Chat context reset.`);
    }
    return true;
}

function removeLibraryDocument(docId, { announce = true } = {}) {
    const targetId = Number(docId);
    const before = libraryDocuments.length;
    libraryDocuments = libraryDocuments.filter(doc => doc.id !== targetId);
    if (libraryDocuments.length === before) return false;
    refreshLibraryPanel();
    resetChatHistoryForLibraryChange('Library updated');
    if (announce) {
        addMsg('system', 'Reference library updated. Chat context reset.');
    }
    return true;
}

function clearLibraryDocuments({ announce = true } = {}) {
    if (libraryDocuments.length === 0) return;
    libraryDocuments = [];
    refreshLibraryPanel();
    resetChatHistoryForLibraryChange('Library cleared');
    if (announce) {
        addMsg('system', 'Reference library cleared. Chat context reset.');
    }
}

// ── XML Helpers (unchanged from original demo) ─────────
function parseXmlStrict(xmlText, label) {
    return parseXmlStrictStandalone(xmlText, label);
}

function getBodyElement(xmlDoc) {
    return getBodyElementFromDocument(xmlDoc);
}

function normalizeBodySectionOrder(xmlDoc) {
    normalizeBodySectionOrderStandalone(xmlDoc);
}

function sanitizeNestedParagraphs(xmlDoc) {
    sanitizeNestedParagraphsInTables(xmlDoc, {
        onInfo: message => log(message)
    });
}

// ── Paragraph helpers ──────────────────────────────────
function getParagraphText(paragraph) {
    return getParagraphTextFromOxml(paragraph);
}

function findParagraphByExactText(xmlDoc, targetText) {
    return findParagraphByBestTextMatch(xmlDoc, targetText, {
        onInfo: message => log(message)
    });
}

function parseParagraphReference(rawValue) {
    return parseParagraphReferenceShared(rawValue);
}

function stripLeadingParagraphMarker(text) {
    return stripLeadingParagraphMarkerShared(text);
}

function splitLeadingParagraphMarker(text) {
    return splitLeadingParagraphMarkerShared(text);
}

function createSimpleParagraph(xmlDoc, text) {
    const p = xmlDoc.createElementNS(NS_W, 'w:p');
    const r = xmlDoc.createElementNS(NS_W, 'w:r');
    const t = xmlDoc.createElementNS(NS_W, 'w:t');
    t.textContent = text;
    r.appendChild(t);
    p.appendChild(r);
    return p;
}

async function runOperation(documentXml, op, author, runtimeContext = null, options = {}) {
    return applyOperationToDocumentXml(documentXml, op, author, runtimeContext, {
        ...options,
        onInfo: message => log(message),
        onWarn: message => log(message)
    });
}

// ── Package artifact helpers ───────────────────────────
async function createNumberingIdState(zip) {
    const existing = await zip.file('word/numbering.xml')?.async('string');
    return createDynamicNumberingIdState(existing || '', {
        minId: 1,
        maxPreferred: 32767
    });
}

function mergeNumberingXml(existingNumberingXml, incomingNumberingXml) {
    return mergeNumberingXmlBySchemaOrder(existingNumberingXml, incomingNumberingXml);
}

async function ensureNumberingArtifacts(zip, numberingXmlList) {
    await ensureNumberingArtifactsInZip(zip, numberingXmlList, {
        mergeNumberingXml: (existingXml, incomingXml) => mergeNumberingXml(existingXml, incomingXml),
        onInfo: message => log(message)
    });
}

async function ensureCommentsArtifacts(zip, commentsXml) {
    await ensureCommentsArtifactsInZip(zip, commentsXml, {
        onInfo: message => log(message)
    });
}

// ── Validation ─────────────────────────────────────────
async function validateOutputDocx(zip) {
    await validateDocxPackage(zip);
}

// ── Logger wiring ──────────────────────────────────────
configureLogger({
    log: (...args) => log(args.map(String).join(' ')),
    warn: (...args) => log(`[WARN] ${args.map(String).join(' ')}`),
    error: (...args) => log(`[ERROR] ${args.map(String).join(' ')}`)
});

// ══════════════════════════════════════════════════════
// ── NEW: Document Ingestion ──────────────────────────
// ══════════════════════════════════════════════════════

async function extractParagraphsFromZip(zip, sourceLabel = 'word/document.xml') {
    const documentXml = await zip.file('word/document.xml')?.async('string');
    if (!documentXml) throw new Error('word/document.xml not found');
    const xmlDoc = parseXmlStrict(documentXml, sourceLabel);
    const body = getBodyElement(xmlDoc);
    if (!body) throw new Error('No w:body in document');

    const paragraphs = [];
    const allP = body.getElementsByTagNameNS(NS_W, 'p');
    for (let i = 0; i < allP.length; i++) {
        const text = getParagraphText(allP[i]).trim();
        if (text) paragraphs.push({ index: i + 1, text });
    }
    return paragraphs;
}

async function extractDocumentParagraphs(zip) {
    return extractParagraphsFromZip(zip, 'word/document.xml');
}

async function extractLibraryDocumentFromFile(file) {
    if (!isLikelyDocxFile(file)) {
        throw new Error(`Unsupported file type for "${file?.name || 'file'}"`);
    }

    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const paragraphs = await extractParagraphsFromZip(zip, `${file.name}:word/document.xml`);
    const fullText = paragraphs.map(p => p.text).join('\n');
    const truncated = truncatePlainText(fullText, LIBRARY_MAX_DOC_CHARS);
    return {
        id: nextLibraryDocId++,
        name: String(file.name || `library-${Date.now()}.docx`),
        size: Number(file.size) || 0,
        paragraphCount: paragraphs.length,
        text: truncated.text,
        originalTextLength: fullText.length,
        truncated: truncated.truncated,
        selected: true
    };
}

function buildLibraryContextLines(libraryDocs) {
    const docs = Array.isArray(libraryDocs) ? libraryDocs : [];
    if (docs.length === 0) return [];

    const lines = [
        'REFERENCE LIBRARY (supplemental context from user-provided documents):'
    ];

    let usedChars = 0;
    let includedDocs = 0;
    for (const doc of docs) {
        const remaining = LIBRARY_MAX_PROMPT_CHARS - usedChars;
        if (remaining <= 0) break;

        const baseText = String(doc.text || '');
        if (!baseText.trim()) continue;
        const take = baseText.length > remaining ? baseText.slice(0, remaining) : baseText;
        const fullyIncluded = take.length === baseText.length;
        includedDocs += 1;
        usedChars += take.length;

        lines.push('');
        lines.push(`[LIB${includedDocs}] ${doc.name} (${doc.paragraphCount} paragraphs, ${doc.originalTextLength} chars source)`);
        lines.push(take);
        if (!fullyIncluded) {
            lines.push('[This library document was truncated to fit prompt limits.]');
            break;
        }
    }

    lines.push('');
    lines.push('Use these library docs as additional background context. Prioritize the uploaded working document when conflicts exist.');
    return lines;
}

// ══════════════════════════════════════════════════════
// ── NEW: Gemini Chat Engine ──────────────────────────
// ══════════════════════════════════════════════════════

function buildSystemInstruction(paragraphs, editModeValue = EDIT_MODE.REDLINE, libraryDocs = []) {
    const normalizedEditMode = normalizeEditMode(editModeValue);
    const directEditsMode = normalizedEditMode === EDIT_MODE.DIRECT;
    const listing = paragraphs.map(p => `[P${p.index}] ${p.text}`).join('\n');
    const libraryContextLines = buildLibraryContextLines(libraryDocs);
    return [
        'You are a contract review AI assistant. The user has uploaded a document.',
        `EDIT MODE: ${directEditsMode ? 'direct edits (apply changes directly, no tracked insert/delete markup)' : 'redlines (tracked changes)'}.`,
        'Below is the document content. Each line is ONE SEPARATE PARAGRAPH, prefixed with [P#]:',
        '',
        listing,
        ...(libraryContextLines.length > 0 ? ['', ...libraryContextLines] : []),
        '',
        'Your job is to analyze the document and perform the operations the user asks for.',
        'For each issue you find, produce an operation. Respond in TWO parts separated by a line that says exactly "---OPERATIONS---":',
        '',
        'PART 1: A conversational explanation of your findings in plain text (what you found, why it matters).',
        '',
        'PART 2: A JSON array of operations. Each operation is one of:',
        '',
        '  { "type": "comment", "targetRef": "P12", "target": "<exact paragraph text>", "textToComment": "<substring to anchor on>", "commentContent": "<your comment>" }',
        '  { "type": "highlight", "targetRef": "P12", "target": "<exact paragraph text>", "textToHighlight": "<substring to highlight>", "color": "yellow|green|cyan|magenta|blue|red" }',
        '  { "type": "redline", "targetRef": "P12", "target": "<exact paragraph text>", "modified": "<replacement paragraph text>" }',
        '  { "type": "redline", "targetRef": "P12", "targetEndRef": "P15", "target": "<exact START paragraph text>", "modified": "<replacement text for P12..P15>" }',
        '',
        'CRITICAL TARGETING RULES:',
        '- Each [P#] line above is a SEPARATE paragraph in the document.',
        '- Always include "targetRef" using the paragraph label (example: "P12").',
        '- "targetRef" must point to the same paragraph as "target".',
        '- "target" MUST be the EXACT text of ONE SINGLE [P#] paragraph. Copy it character-for-character.',
        '- NEVER include the [P#] prefix in ANY operation field. The [P#] prefix is only a reference label, NOT part of the actual text.',
        '- NEVER combine or concatenate text from multiple [P#] paragraphs into one target.',
        '- If you need to modify multiple paragraphs, create a SEPARATE operation for EACH paragraph.',
        '- EXCEPTION: for structural conversions (especially text->table), use ONE redline with "targetRef" as start and "targetEndRef" as end of the contiguous block.',
        '- "textToComment" / "textToHighlight" must be an exact substring found within that single paragraph.',
        '',
        'OPERATION RULES:',
        '- Use "comment" to explain issues (best for deviations from market standards).',
        '- Use "highlight" to draw visual attention to problematic phrases.',
        '- Use "redline" to suggest replacement language for a single paragraph.',
        ...(directEditsMode
            ? [
                '- In direct-edit mode, prefer "redline" operations for actual language changes.',
                '- Unless the user explicitly asks for annotations, avoid comment/highlight-only output.'
            ]
            : []),
        '',
        'FORMATTING IN REDLINES:',
        '- The "modified" field in redline operations supports special formatting syntax:',
        '  - **bold text** → wraps text in bold (use double asterisks)',
        '  - ++underline text++ → wraps text in underline (use double plus signs)',
        '  - Bullet lists: start each line with "- " for top-level bullets, "  - " for nested bullets',
        '  - Ordered lists: use exactly one marker per line (examples: "1. ...", "A. ...", "a. ...", "I. ...", "i. ...")',
        '  - NEVER double-mark a list item (invalid: "- A. ...", "1. a. ..."). Use only the final desired marker.',
        '  - For ordered lists, do not include marker characters inside item text.',
        '  - Tables: use markdown table syntax (e.g., "| Col1 | Col2 |\\n|---|---|\\n| val | val |")',
        '- When converting multiple paragraphs into a table, you MUST set "targetEndRef" to include the full source block.',
        '- For EXISTING TABLE STRUCTURE changes (add/remove/reorder rows/columns), the "modified" value MUST be the FULL markdown table for that target table, not a single cell value.',
        '- For table structure changes, target any paragraph within that table and include the correct "targetRef".',
        '- For one table structure request, return ONE operation for that table. Do not emit parallel per-cell/per-column duplicates for the same row insertion.',
        '- If you can only express it as multiline cell text (example: "Title:\\nDate:"), the client may convert it to full table markdown automatically, but returning full markdown table is preferred.',
        '- For list insertion in the middle of an existing list, return list markdown that includes the existing target item followed by inserted item(s), each on its own list line.',
        '- If the target paragraph is in an ordered list, preserve ordered markers for inserted lines (for example use "2.2.1."), not bullet markers.',
        '- You CAN apply formatting like bold and underline using redline operations.',
        '- To underline a title, use: { "type": "redline", "target": "Title Text", "modified": "++Title Text++" }',
        '- To bold a word, use: { "type": "redline", "target": "Some text here", "modified": "Some **text** here" }',
        '- To add NEW content before an existing paragraph, use a redline that prepends the new text before the original.',
        '- You may return an empty array [] if there are no issues.',
        '- Keep comments concise and actionable.',
        '- If the user asks about "market standards", focus on: unusual liability caps, atypical indemnification, non-standard termination, unreasonable non-compete, missing limitation of liability, unusual governing law, missing confidentiality, unusual assignment restrictions, non-standard warranty disclaimers, missing force majeure.',
        '- Prefer "comment" operations for explanations and "redline" for suggesting replacement language.',
    ].join('\n');
}

function parseGeminiChatResponse(rawText) {
    const separatorIdx = rawText.indexOf('---OPERATIONS---');
    let explanationText;
    let operationsJson;

    if (separatorIdx >= 0) {
        explanationText = rawText.slice(0, separatorIdx).trim();
        operationsJson = rawText.slice(separatorIdx + '---OPERATIONS---'.length).trim();
    } else {
        // Try to find a JSON array anywhere in the response
        const jsonMatch = rawText.match(/\[[\s\S]*\]/);
        if (jsonMatch) {
            const jsonStart = rawText.indexOf(jsonMatch[0]);
            explanationText = rawText.slice(0, jsonStart).trim();
            operationsJson = jsonMatch[0];
        } else {
            return { explanation: rawText, operations: [] };
        }
    }

    // Strip markdown fences if present
    operationsJson = operationsJson.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();

    let operations = [];
    try {
        // Find the JSON array within the text
        const arrayMatch = operationsJson.match(/\[[\s\S]*\]/);
        if (arrayMatch) {
            operations = JSON.parse(arrayMatch[0]);
        }
    } catch (err) {
        log(`[WARN] Could not parse operations JSON: ${err.message}`);
    }

    // Validate and normalize each operation
    operations = operations
        .map(op => {
            if (!op || typeof op !== 'object') return null;
            const type = String(op.type || '').toLowerCase().trim();
            if (!type) return null;

            const splitTarget = splitLeadingParagraphMarker(op.target);
            const explicitRef = parseParagraphReference(op.targetRef ?? op.paragraphRef ?? op.paragraphIndex ?? op.targetIndex);
            const explicitEndRef = parseParagraphReference(op.targetEndRef ?? op.endTargetRef ?? op.endParagraphRef ?? op.endParagraphIndex);
            const targetRef = explicitRef || splitTarget.targetRef || null;
            const target = splitTarget.text;

            const normalizedOp = { ...op, type, target, targetRef, targetEndRef: explicitEndRef || null };
            if (normalizedOp.modified != null) normalizedOp.modified = stripLeadingParagraphMarker(normalizedOp.modified);
            if (normalizedOp.textToComment != null) normalizedOp.textToComment = stripLeadingParagraphMarker(normalizedOp.textToComment);
            if (normalizedOp.textToHighlight != null) normalizedOp.textToHighlight = stripLeadingParagraphMarker(normalizedOp.textToHighlight);
            if (normalizedOp.commentContent != null) normalizedOp.commentContent = String(normalizedOp.commentContent).trim();

            if (type === 'highlight') {
                const c = String(normalizedOp.color || '').toLowerCase();
                normalizedOp.color = ALLOWED_HIGHLIGHT_COLORS.includes(c) ? c : 'yellow';
            }

            return normalizedOp;
        })
        .filter(op => {
            if (!op || !op.type) return false;
            if (!op.target && !op.targetRef) return false;
            if (op.type === 'comment' && (!op.textToComment || !op.commentContent)) return false;
            if (op.type === 'highlight' && !op.textToHighlight) return false;
            if (op.type === 'redline' && !op.modified) return false;
            return op.type === 'comment' || op.type === 'highlight' || op.type === 'redline';
        });

    return { explanation: explanationText || '(No explanation provided)', operations };
}

function maskGeminiEndpoint(endpoint) {
    return String(endpoint || '').replace(/([?&]key=)[^&]+/i, '$1***');
}

function buildGeminiRequestPayload(userMessage, paragraphs, editModeValue, libraryDocs, apiKey) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(apiKey)}`;
    const systemInstruction = buildSystemInstruction(paragraphs, editModeValue, libraryDocs);
    const contents = [
        ...chatHistory,
        { role: 'user', parts: [{ text: userMessage }] }
    ];
    const requestBody = {
        systemInstruction: { parts: [{ text: systemInstruction }] },
        contents,
        generationConfig: {
            temperature: 0.3,
            maxOutputTokens: 8000
        }
    };
    return { endpoint, requestBody };
}

function captureGeminiRequestDebug(endpoint, requestBody, selectedLibraryDocsCount, totalLibraryDocCount) {
    const systemText = String(requestBody?.systemInstruction?.parts?.[0]?.text || '');
    const systemInstructionPreview = systemText.length > GEMINI_REQUEST_PREVIEW_CHARS
        ? `${systemText.slice(0, GEMINI_REQUEST_PREVIEW_CHARS)}\n...[truncated in preview]`
        : systemText;
    const debugPayload = {
        method: 'POST',
        endpoint: maskGeminiEndpoint(endpoint),
        headers: { 'Content-Type': 'application/json' },
        body: requestBody,
        meta: {
            selectedLibraryDocsCount,
            totalLibraryDocCount
        },
        systemInstructionPreview
    };
    lastGeminiRequestDebug = debugPayload;
    if (typeof window !== 'undefined') {
        window.__BROWSER_DEMO_LAST_GEMINI_REQUEST__ = debugPayload;
    }
    log(`[Gemini] Request built (${selectedLibraryDocsCount}/${totalLibraryDocCount} library docs selected). Inspect window.__BROWSER_DEMO_LAST_GEMINI_REQUEST__ in devtools.`);
}

async function sendGeminiChat(userMessage, paragraphs, apiKey, editModeValue = EDIT_MODE.REDLINE, libraryDocs = [], options = {}) {
    const selectedLibraryDocsCount = Array.isArray(libraryDocs) ? libraryDocs.length : 0;
    const totalLibraryDocCount = Number.isInteger(options?.totalLibraryDocCount) ? options.totalLibraryDocCount : selectedLibraryDocsCount;
    const { endpoint, requestBody } = buildGeminiRequestPayload(userMessage, paragraphs, editModeValue, libraryDocs, apiKey);
    captureGeminiRequestDebug(endpoint, requestBody, selectedLibraryDocsCount, totalLibraryDocCount);

    const response = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Gemini API ${response.status}: ${errorText.slice(0, 400)}`);
    }

    const payload = await response.json();
    const rawText = payload?.candidates?.[0]?.content?.parts?.map(p => p?.text || '').join('').trim() || '';
    if (!rawText) throw new Error('Gemini returned empty response');

    // Update chat history for multi-turn
    chatHistory.push({ role: 'user', parts: [{ text: userMessage }] });
    chatHistory.push({ role: 'model', parts: [{ text: rawText }] });

    return parseGeminiChatResponse(rawText);
}

// ══════════════════════════════════════════════════════
// ── NEW: Apply Chat Operations to Document ──────────
// ══════════════════════════════════════════════════════

async function applyChatOperations(zip, operations, author, editModeValue = EDIT_MODE.REDLINE) {
    const normalizedEditMode = normalizeEditMode(editModeValue);
    const generateRedlines = shouldGenerateRedlines(normalizedEditMode);
    let documentXml = await zip.file('word/document.xml')?.async('string');
    if (!documentXml) throw new Error('word/document.xml not found');
    parseXmlStrict(documentXml, 'word/document.xml');

    const capturedNumberingXml = [];
    const capturedCommentsXml = [];
    const results = [];
    const snapshotDoc = parseXmlStrict(documentXml, 'word/document.xml (target snapshot)');
    const runtimeContext = {
        numberingIdState: await createNumberingIdState(zip),
        targetRefSnapshot: buildTargetReferenceSnapshot(snapshotDoc),
        tableStructuralRedlineKeys: new Set(),
        listFallbackSharedNumIdByKey: new Map(),
        listFallbackSequenceState: {
            explicitByNumberingKey: new Map()
        }
    };

    for (const op of operations) {
        const targetRefLabel = op.targetRef
            ? (op.targetEndRef ? `[P${op.targetRef}-P${op.targetEndRef}] ` : `[P${op.targetRef}] `)
            : '';
        const label = `${op.type}: ${targetRefLabel}"${(op.target || '').slice(0, 50)}…"`;
        log(`Applying: ${label}`);
        try {
            const step = await runOperation(documentXml, op, author, runtimeContext, { generateRedlines });
            documentXml = step.documentXml;
            if (step.numberingXml) capturedNumberingXml.push(step.numberingXml);
            if (step.commentsXml) capturedCommentsXml.push(step.commentsXml);
            if (step.warnings?.length > 0) {
                for (const warning of step.warnings) log(`  warning: ${warning}`);
            }
            results.push({ ...op, success: step.hasChanges, error: null });
            log(`  → ${step.hasChanges ? 'applied' : 'no change'}`);
        } catch (err) {
            const errorMsg = err?.message || String(err);
            log(`  → FAILED: ${errorMsg}`);
            results.push({ ...op, success: false, error: errorMsg });
        }
    }

    // Normalize, sanitize, and write back
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const finalDoc = parser.parseFromString(documentXml, 'application/xml');
    normalizeBodySectionOrder(finalDoc);
    sanitizeNestedParagraphs(finalDoc);
    documentXml = serializer.serializeToString(finalDoc);
    zip.file('word/document.xml', documentXml);

    await ensureNumberingArtifacts(zip, capturedNumberingXml);
    for (const cx of capturedCommentsXml) await ensureCommentsArtifacts(zip, cx);

    try {
        await validateOutputDocx(zip);
    } catch (validationErr) {
        log(`[WARN] Post-operation validation: ${validationErr.message}`);
        // Non-fatal — document may still be usable
    }

    return results;
}

// ── Operation summary HTML builder ─────────────────────
function buildOpSummaryHtml(results) {
    if (!results || results.length === 0) return '';
    const items = results.map(r => {
        const badge = `<span class="op-badge ${r.type}">${r.type}</span>`;
        const status = r.success ? '✓' : (r.error ? `✗ ${r.error.slice(0, 60)}` : '—');
        const target = (r.target || '').slice(0, 60);
        return `<div class="op-item">${badge} <span title="${escapeHtml(r.target || '')}">${escapeHtml(target)}…</span> <span style="margin-left:auto;opacity:0.7">${status}</span></div>`;
    }).join('');
    const applied = results.filter(r => r.success).length;
    return `<div class="op-summary"><strong>${applied}/${results.length} operations applied</strong>${items}</div>`;
}

function escapeHtml(str) {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── Download ───────────────────────────────────────────
function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ══════════════════════════════════════════════════════
// ── EVENT WIRING ─────────────────────────────────────
// ══════════════════════════════════════════════════════

function setLibraryControlsDisabled(disabled) {
    if (libraryAddBtn) libraryAddBtn.disabled = disabled;
    if (libraryClearBtn) libraryClearBtn.disabled = disabled;
    if (libraryDocxFilesInput) libraryDocxFilesInput.disabled = disabled;
}

async function ingestLibraryFiles(fileList, sourceLabel = 'library') {
    const candidates = Array.from(fileList || []).filter(Boolean);
    if (candidates.length === 0) return;

    const docxFiles = candidates.filter(isLikelyDocxFile);
    if (docxFiles.length === 0) {
        addMsg('system warn', 'No valid .docx files found for the reference library.');
        return;
    }

    setLibraryControlsDisabled(true);
    addMsg('system', `Loading ${docxFiles.length} library document(s) from ${escapeHtml(sourceLabel)}…`);

    let added = 0;
    let replaced = 0;
    let failed = 0;
    const failures = [];

    try {
        for (const file of docxFiles) {
            try {
                const parsed = await extractLibraryDocumentFromFile(file);
                const existingIdx = libraryDocuments.findIndex(
                    doc => String(doc.name || '').toLowerCase() === parsed.name.toLowerCase()
                );
                if (existingIdx >= 0) {
                    parsed.selected = libraryDocuments[existingIdx].selected !== false;
                    parsed.id = libraryDocuments[existingIdx].id;
                    libraryDocuments.splice(existingIdx, 1, parsed);
                    replaced += 1;
                } else {
                    libraryDocuments.push(parsed);
                    added += 1;
                }
                log(`[Library] Loaded ${parsed.name} (${parsed.paragraphCount} paragraphs, ${parsed.text.length} chars${parsed.truncated ? ', truncated' : ''})`);
            } catch (error) {
                failed += 1;
                const msg = error?.message || String(error);
                failures.push(`${file.name}: ${msg}`);
                log(`[WARN] [Library] Failed to load ${file.name}: ${msg}`);
            }
        }
    } finally {
        setLibraryControlsDisabled(false);
    }

    libraryDocuments.sort((a, b) => String(a.name || '').localeCompare(String(b.name || ''), undefined, { sensitivity: 'base' }));
    refreshLibraryPanel();

    const hasChanges = added > 0 || replaced > 0;
    if (hasChanges) {
        resetChatHistoryForLibraryChange('Reference library updated');
    }

    if (hasChanges) {
        const parts = [];
        if (added > 0) parts.push(`${added} added`);
        if (replaced > 0) parts.push(`${replaced} replaced`);
        if (failed > 0) parts.push(`${failed} failed`);
        addMsg('system success', `Reference library updated: ${parts.join(', ')}.${hasChanges ? ' Chat context reset.' : ''}`);
    } else if (failed > 0) {
        addMsg('system error', 'Failed to add library documents. Check engine log for details.');
    }

    if (failures.length > 0) {
        log(`[Library] Failures:\n- ${failures.join('\n- ')}`);
    }
}

// File upload → load document, extract paragraphs
fileInput.addEventListener('change', async () => {
    const file = fileInput.files?.[0];
    if (!file) return;

    try {
        addMsg('system', `Loading <strong>${escapeHtml(file.name)}</strong>…`);
        currentZip = await JSZip.loadAsync(await file.arrayBuffer());
        documentParagraphs = await extractDocumentParagraphs(currentZip);
        chatHistory = [];
        operationCount = 0;
        downloadBtn.style.display = 'none';
        downloadXmlBtn.style.display = '';
        downloadXmlBtn.disabled = false;

        addMsg('system success', `Document loaded: <strong>${documentParagraphs.length} paragraphs</strong> found. You can now ask the AI to review it.`);
        log(`[Demo] v${DEMO_VERSION} — loaded ${file.name} (${documentParagraphs.length} paragraphs)`);

        chatInput.disabled = false;
        sendBtn.disabled = false;
        await renderPreviewFromZip(currentZip, 'Upload');
        chatInput.focus();
    } catch (err) {
        addMsg('system error', `Failed to load document: ${escapeHtml(err.message || String(err))}`);
        log(`[FATAL] ${err.message || String(err)}`);
    }
});

// Send chat message
async function handleSend() {
    const userText = chatInput.value.trim();
    if (!userText) return;
    const apiKey = geminiApiKeyInput?.value.trim() || getStoredGeminiApiKey();
    if (!apiKey) {
        addMsg('system warn', 'Please enter and save your Gemini API key first.');
        return;
    }
    if (!currentZip) {
        addMsg('system warn', 'Please upload a .docx file first.');
        return;
    }

    addMsg('user', escapeHtml(userText));
    chatInput.value = '';
    chatInput.style.height = 'auto';
    sendBtn.disabled = true;
    chatInput.disabled = true;

    const thinkingEl = addMsg('system', '⏳ Analyzing document…');
    const selectedLibraryDocs = getSelectedLibraryDocuments();

    try {
        const result = await sendGeminiChat(userText, documentParagraphs, apiKey, editMode, selectedLibraryDocs, {
            totalLibraryDocCount: libraryDocuments.length
        });

        // Remove thinking indicator
        thinkingEl.remove();

        let assistantHtml = escapeHtml(result.explanation).replace(/\n/g, '<br>');

        if (result.operations.length > 0) {
            addMsg('system', `Applying ${result.operations.length} operation(s) in <strong>${getEditModeLabel()}</strong> mode…`);
            const author = authorInput.value.trim() || 'Browser Demo AI';
            const opResults = await applyChatOperations(currentZip, result.operations, author, editMode);
            operationCount += opResults.filter(r => r.success).length;

            // Re-extract paragraphs after modifications
            documentParagraphs = await extractDocumentParagraphs(currentZip);
            await renderPreviewFromZip(currentZip, 'Chat');

            assistantHtml += buildOpSummaryHtml(opResults);

            // Show download button
            downloadBtn.style.display = '';
            downloadBtn.disabled = false;
            downloadXmlBtn.style.display = '';
            downloadXmlBtn.disabled = false;
        } else {
            assistantHtml += '<div class="op-summary" style="color:var(--muted)">No document operations returned.</div>';
        }

        addMsg('assistant', assistantHtml);
        log(`[Chat] Turn complete — ${result.operations.length} ops returned, ${operationCount} total applied`);

    } catch (err) {
        thinkingEl.remove();
        addMsg('system error', `Error: ${escapeHtml(err.message || String(err))}`);
        log(`[ERROR] ${err.message || String(err)}`);
        console.error(err);
    } finally {
        sendBtn.disabled = !currentZip;
        chatInput.disabled = !currentZip;
        chatInput.focus();
    }
}

sendBtn.addEventListener('click', handleSend);
chatInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        handleSend();
    }
});

// Auto-resize textarea
chatInput.addEventListener('input', () => {
    chatInput.style.height = 'auto';
    chatInput.style.height = Math.min(chatInput.scrollHeight, 120) + 'px';
});

if (libraryAddBtn && libraryDocxFilesInput) {
    libraryAddBtn.addEventListener('click', () => libraryDocxFilesInput.click());
    libraryDocxFilesInput.addEventListener('change', async () => {
        try {
            await ingestLibraryFiles(libraryDocxFilesInput.files, 'file picker');
        } finally {
            libraryDocxFilesInput.value = '';
        }
    });
}

if (libraryToggleBtn) {
    libraryToggleBtn.addEventListener('click', () => {
        setLibraryCollapsed(!isLibraryCollapsed, { persist: true });
    });
}

if (libraryClearBtn) {
    libraryClearBtn.addEventListener('click', () => {
        clearLibraryDocuments({ announce: true });
    });
}

if (libraryItemsEl) {
    libraryItemsEl.addEventListener('change', (event) => {
        const toggle = event.target?.closest?.('input[data-library-doc-toggle-id]');
        if (!toggle) return;
        setLibraryDocumentSelected(toggle.dataset.libraryDocToggleId, toggle.checked, { announce: false });
    });

    libraryItemsEl.addEventListener('click', (event) => {
        const button = event.target?.closest?.('button[data-library-doc-id]');
        if (!button) return;
        removeLibraryDocument(button.dataset.libraryDocId, { announce: true });
    });
}

if (libraryDropZone) {
    const preventDefault = (event) => {
        event.preventDefault();
        event.stopPropagation();
    };
    const setDragState = (active) => {
        libraryDropZone.classList.toggle('drag-over', !!active);
    };

    ['dragenter', 'dragover'].forEach(type => {
        libraryDropZone.addEventListener(type, (event) => {
            preventDefault(event);
            setDragState(true);
        });
    });
    ['dragleave', 'dragend'].forEach(type => {
        libraryDropZone.addEventListener(type, (event) => {
            preventDefault(event);
            setDragState(false);
        });
    });
    libraryDropZone.addEventListener('drop', async (event) => {
        preventDefault(event);
        setDragState(false);
        const droppedFiles = event.dataTransfer?.files;
        await ingestLibraryFiles(droppedFiles, 'drag and drop');
    });
}

// Download button
downloadBtn.addEventListener('click', async () => {
    if (!currentZip) return;
    try {
        const blob = await currentZip.generateAsync({ type: 'blob' });
        const originalName = fileInput.files?.[0]?.name || 'document.docx';
        const outputName = originalName.replace(/\.docx$/i, '') + '-reviewed.docx';
        downloadBlob(blob, outputName);
        addMsg('system success', `Downloaded <strong>${escapeHtml(outputName)}</strong>`);
    } catch (err) {
        addMsg('system error', `Download failed: ${escapeHtml(err.message || String(err))}`);
    }
});

// Debug XML download button (word/document.xml from current in-memory package)
downloadXmlBtn.addEventListener('click', async () => {
    if (!currentZip) return;
    try {
        const documentXml = await currentZip.file('word/document.xml')?.async('string');
        if (!documentXml) throw new Error('word/document.xml not found in current package');

        const xmlBlob = new Blob([documentXml], { type: 'application/xml;charset=utf-8' });
        const originalName = fileInput.files?.[0]?.name || 'document.docx';
        const outputName = originalName.replace(/\.docx$/i, '') + '-document.xml';
        downloadBlob(xmlBlob, outputName);
        addMsg('system success', `Downloaded <strong>${escapeHtml(outputName)}</strong>`);
    } catch (err) {
        addMsg('system error', `XML download failed: ${escapeHtml(err.message || String(err))}`);
    }
});

// Log panel toggle
logToggle.addEventListener('click', () => {
    const isVisible = logEl.style.display === 'block';
    logEl.style.display = isVisible ? 'none' : 'block';
    logToggle.textContent = isVisible ? '▶ Engine log' : '▼ Engine log';
});

if (refreshPreviewBtn) {
    refreshPreviewBtn.addEventListener('click', async () => {
        if (!currentZip) {
            setPreviewStatus('Upload a .docx file before refreshing preview.');
            return;
        }
        await renderPreviewFromZip(currentZip, 'Manual Refresh');
    });
}

// Save Gemini API key
saveGeminiKeyBtn.addEventListener('click', () => {
    const key = geminiApiKeyInput?.value.trim() || '';
    const saved = setStoredGeminiApiKey(key);
    if (!saved) { addMsg('system warn', 'Unable to save API key in this browser.'); return; }
    if (key) addMsg('system success', 'Gemini API key saved.');
    else addMsg('system', 'Gemini API key cleared.');
});

// Restore saved key
if (geminiApiKeyInput) {
    const storedKey = getStoredGeminiApiKey();
    if (storedKey) geminiApiKeyInput.value = storedKey;
}

refreshLibraryPanel();
setLibraryCollapsed(isLibraryCollapsed, { persist: false });

// Restore + wire edit mode toggle
setEditMode(editMode, { announce: false, resetChatHistory: false });
for (const input of editModeInputs) {
    input?.addEventListener('change', () => {
        if (!input.checked) return;
        setEditMode(input.value, { announce: true, resetChatHistory: true });
    });
}

// ══════════════════════════════════════════════════════
// ── LEGACY: Kitchen-Sink Demo (preserved) ────────────
// ══════════════════════════════════════════════════════

function ensureDemoTargets(documentXml) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');
    const body = getBodyElement(xmlDoc);
    if (!body) throw new Error('Could not find w:body');
    for (const marker of DEMO_MARKERS) {
        if (!findParagraphByExactText(xmlDoc, marker)) {
            insertBodyElementBeforeSectPr(body, createSimpleParagraph(xmlDoc, marker));
            log(`Inserted missing marker: ${marker}`);
        }
    }
    normalizeBodySectionOrder(xmlDoc);
    return serializer.serializeToString(xmlDoc);
}

function defaultTokenForTarget(target) {
    if (target === 'DEMO FORMAT TARGET') return 'FORMAT';
    if (target === 'DEMO_TEXT_TARGET') return 'DEMO_TEXT_TARGET';
    if (target === 'DEMO_LIST_TARGET') return 'Browser';
    return 'Status';
}

function normalizeGeminiToolAction(rawAction) {
    const tool = String(rawAction?.tool || '').toLowerCase().trim();
    const args = rawAction?.args && typeof rawAction.args === 'object' ? rawAction.args : {};
    const rawTarget = String(args.target || '').trim();
    const target = DEMO_MARKERS.includes(rawTarget) ? rawTarget : 'DEMO FORMAT TARGET';
    if (tool === 'comment') {
        return { type: 'comment', label: 'Gemini Surprise Tool Action', target, textToComment: String(args.textToComment || defaultTokenForTarget(target)).trim() || defaultTokenForTarget(target), commentContent: (String(args.commentContent || 'Gemini surprise comment.').trim() || 'Gemini surprise comment.').slice(0, 220) };
    }
    if (tool === 'highlight') {
        const c = String(args.color || '').trim().toLowerCase();
        return { type: 'highlight', label: 'Gemini Surprise Tool Action', target, textToHighlight: String(args.textToHighlight || defaultTokenForTarget(target)).trim() || defaultTokenForTarget(target), color: ALLOWED_HIGHLIGHT_COLORS.includes(c) ? c : 'yellow' };
    }
    if (tool === 'redline') {
        return { type: 'redline', label: 'Gemini Surprise Tool Action', target, modified: (String(args.modified || `${target} refined by Gemini.`).trim() || `${target} refined by Gemini.`).slice(0, 260) };
    }
    throw new Error(`Unsupported Gemini tool: "${tool}"`);
}

function extractJsonObject(text) {
    if (!text) throw new Error('No Gemini tool action text');
    const trimmed = text.trim();
    try { return JSON.parse(trimmed); } catch { }
    const fenced = trimmed.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
    if (fenced?.[1]) return JSON.parse(fenced[1]);
    const f = trimmed.indexOf('{'), l = trimmed.lastIndexOf('}');
    if (f >= 0 && l > f) return JSON.parse(trimmed.slice(f, l + 1));
    throw new Error('No JSON object found');
}

async function generateGeminiRedlineSuggestion(originalText, apiKey) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(apiKey)}`;
    const prompt = ['Rewrite the following text as a cleaner sentence for a professional document.', 'Return plain text only, no quotes, markdown, bullets, or explanation.', `Text: ${originalText}`].join('\n');
    const response = await fetch(endpoint, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }], generationConfig: { temperature: 0.4, maxOutputTokens: 80 } }) });
    if (!response.ok) throw new Error(`Gemini API ${response.status}: ${(await response.text()).slice(0, 300)}`);
    const payload = await response.json();
    const suggestion = (payload?.candidates?.[0]?.content?.parts?.map(p => p?.text || '').join(' ').trim() || '');
    if (!suggestion) throw new Error('Gemini returned no text suggestion');
    return suggestion;
}

async function generateGeminiToolAction(apiKey) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(apiKey)}`;
    const prompt = ['You are choosing one surprise action for a DOCX demo pipeline.', 'Return JSON only, no markdown and no commentary.', `Allowed targets: ${DEMO_MARKERS.join(', ')}`, 'Choose exactly one tool:', '- comment -> args: { "target": string, "textToComment": string, "commentContent": string }', '- highlight -> args: { "target": string, "textToHighlight": string, "color": string }', '- redline -> args: { "target": string, "modified": string }', 'Keep args short.', '{ "tool": "comment|highlight|redline", "args": { ... } }'].join('\n');
    const response = await fetch(endpoint, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }], generationConfig: { temperature: 0.9, maxOutputTokens: 200 } }) });
    if (!response.ok) throw new Error(`Gemini API ${response.status}: ${(await response.text()).slice(0, 300)}`);
    const payload = await response.json();
    const rawText = (payload?.candidates?.[0]?.content?.parts?.map(p => p?.text || '').join('\n').trim() || '');
    if (!rawText) throw new Error('Gemini returned no tool action');
    return normalizeGeminiToolAction(extractJsonObject(rawText));
}

function buildSurpriseFallbackOperation(op) {
    const safeTarget = 'DEMO FORMAT TARGET';
    if (op.type === 'highlight') return { ...op, target: safeTarget, textToHighlight: 'FORMAT', color: ALLOWED_HIGHLIGHT_COLORS.includes(String(op.color || '').toLowerCase()) ? op.color : 'yellow' };
    if (op.type === 'comment') return { ...op, target: safeTarget, textToComment: 'FORMAT' };
    return { ...op, target: safeTarget, modified: 'DEMO FORMAT TARGET updated by Gemini surprise retry.' };
}

async function runKitchenSink(inputFile, author, geminiApiKey) {
    const zip = await JSZip.loadAsync(await inputFile.arrayBuffer());
    const documentFile = zip.file('word/document.xml');
    if (!documentFile) throw new Error('word/document.xml not found');
    let documentXml = await documentFile.async('string');
    parseXmlStrict(documentXml, 'word/document.xml (input)');
    documentXml = ensureDemoTargets(documentXml);

    const capturedNumberingXml = [];
    const capturedCommentsXml = [];
    const runtimeContext = {
        numberingIdState: await createNumberingIdState(zip),
        listFallbackSharedNumIdByKey: new Map(),
        listFallbackSequenceState: {
            explicitByNumberingKey: new Map()
        }
    };
    const fallbackRedlineText = 'DEMO_TEXT_TARGET rewritten with extra words from the browser demo.';
    const fallbackToolOperation = { type: 'comment', label: 'AI Surprise Fallback', target: 'DEMO FORMAT TARGET', textToComment: 'FORMAT', commentContent: 'Fallback AI action: please review the formatting language here.' };
    let geminiRedlineText = fallbackRedlineText;
    let geminiToolOperation = fallbackToolOperation;

    if (geminiApiKey) {
        log('Generating Gemini redline suggestion for DEMO_TEXT_TARGET...');
        try {
            const suggested = await generateGeminiRedlineSuggestion('DEMO_TEXT_TARGET', geminiApiKey);
            if (suggested.trim() && suggested.trim() !== 'DEMO_TEXT_TARGET') { geminiRedlineText = suggested.trim(); log(`Gemini suggestion: ${geminiRedlineText}`); }
            else log('[WARN] Gemini suggestion matched source; using fallback.');
        } catch (error) { log(`[WARN] Gemini suggestion failed; fallback. ${error.message || String(error)}`); }

        log('Generating Gemini surprise tool action...');
        try {
            const action = await generateGeminiToolAction(geminiApiKey);
            geminiToolOperation = action;
            log(`Surprise action: ${action.type} on "${action.target}"`);
        } catch (error) { log(`[WARN] Surprise action failed; fallback. ${error.message || String(error)}`); }
    } else { log('[WARN] No Gemini API key; using fallbacks.'); }

    const operations = [
        { type: 'redline', label: 'Text Edit', target: 'DEMO_TEXT_TARGET', modified: geminiRedlineText },
        { type: 'redline', label: 'Format-Only', target: 'DEMO FORMAT TARGET', modified: '**DEMO** ++FORMAT++ TARGET' },
        { type: 'redline', label: 'Bullets', target: 'DEMO_LIST_TARGET', modified: ['- Browser demo top bullet', '  - Nested bullet A', '  - Nested bullet B', '- Browser demo second bullet'].join('\n') },
        { type: 'redline', label: 'Table', target: 'DEMO_TABLE_TARGET', modified: ['| Item | Owner | Status |', '|---|---|---|', '| Engine refactor | Platform | Done |', '| Browser demo | UX | In Progress |', '| Documentation | QA | Planned |'].join('\n') },
        { ...geminiToolOperation }
    ];

    for (const op of operations) {
        log(`Running: ${op.label}`);
        let step;
        try { step = await runOperation(documentXml, op, author, runtimeContext); }
        catch (error) {
            const msg = error?.message || String(error);
            const isSurprise = op.label === 'Gemini Surprise Tool Action' || op.label === 'AI Surprise Fallback';
            if (!isSurprise || !msg.includes('Target paragraph not found')) throw error;
            log(`[WARN] ${msg}`);
            log('[WARN] Retrying on safe target.');
            step = await runOperation(documentXml, buildSurpriseFallbackOperation(op), author, runtimeContext);
        }
        documentXml = step.documentXml;
        if (step.numberingXml) capturedNumberingXml.push(step.numberingXml);
        if (step.commentsXml) capturedCommentsXml.push(step.commentsXml);
        if (step.warnings?.length > 0) for (const w of step.warnings) log(`  warning: ${w}`);
        log(`  changed: ${step.hasChanges}`);
    }

    { const p = new DOMParser(), s = new XMLSerializer(), d = p.parseFromString(documentXml, 'application/xml'); normalizeBodySectionOrder(d); documentXml = s.serializeToString(d); }
    zip.file('word/document.xml', documentXml);
    await ensureNumberingArtifacts(zip, capturedNumberingXml);
    for (const cx of capturedCommentsXml) await ensureCommentsArtifacts(zip, cx);
    await validateOutputDocx(zip);
    return zip;
}

// Kitchen-sink button
runBtn.addEventListener('click', async () => {
    const file = fileInput.files?.[0];
    if (!file) { addMsg('system warn', 'Please choose a .docx file.'); return; }
    runBtn.disabled = true;
    logEl.textContent = '';
    addMsg('system', 'Running kitchen-sink demo…');
    try {
        const author = authorInput.value.trim() || 'Browser Demo AI';
        const geminiApiKey = geminiApiKeyInput?.value.trim() || '';
        if (geminiApiKey) setStoredGeminiApiKey(geminiApiKey);
        log(`[Demo] Version: ${DEMO_VERSION}`);
        const outputZip = await runKitchenSink(file, author, geminiApiKey);
        const output = await outputZip.generateAsync({ type: 'blob' });
        const outputName = file.name.replace(/\.docx$/i, '') + '-kitchen-sink-demo.docx';
        downloadBlob(output, outputName);
        currentZip = outputZip;
        documentParagraphs = await extractDocumentParagraphs(currentZip);
        chatHistory = [];
        operationCount = 0;
        sendBtn.disabled = false;
        chatInput.disabled = false;
        downloadBtn.style.display = '';
        downloadBtn.disabled = false;
        downloadXmlBtn.style.display = '';
        downloadXmlBtn.disabled = false;
        await renderPreviewFromZip(currentZip, 'Kitchen Sink');
        addMsg('system success', 'Kitchen-sink demo completed. Document downloaded.');
        log('Kitchen-sink demo completed successfully.');
    } catch (err) {
        addMsg('system error', `Kitchen-sink failed: ${escapeHtml(err.message || String(err))}`);
        log(`[FATAL] ${err.message || String(err)}`);
        console.error(err);
    } finally { runBtn.disabled = false; }
});
