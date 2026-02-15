import JSZip from 'https://esm.sh/jszip@3.10.1';
import {
    applyRedlineToOxml,
    reconcileMarkdownTableOoxml,
    applyHighlightToOoxml,
    injectCommentsIntoOoxml,
    configureLogger,
    getParagraphText as getParagraphTextFromOxml,
    buildTargetReferenceSnapshot,
    isMarkdownTableText,
    findContainingWordElement,
    findParagraphByStrictText as findParagraphByStrictTextShared,
    findParagraphByBestTextMatch,
    parseParagraphReference as parseParagraphReferenceShared,
    stripLeadingParagraphMarker as stripLeadingParagraphMarkerShared,
    splitLeadingParagraphMarker as splitLeadingParagraphMarkerShared,
    resolveTargetParagraphWithSnapshot as resolveTargetParagraphWithSnapshotShared,
    buildSingleLineListStructuralFallbackPlan,
    executeSingleLineListStructuralFallback,
    resolveSingleLineListFallbackNumberingAction,
    recordSingleLineListFallbackExplicitSequence,
    clearSingleLineListFallbackExplicitSequence,
    enforceListBindingOnParagraphNodes,
    synthesizeTableMarkdownFromMultilineCellEdit,
    synthesizeExpandedListScopeEdit,
    planListInsertionOnlyEdit,
    getParagraphListInfo,
    stripRedundantLeadingListMarkers,
    stripSingleLineListMarkerPrefix,
    createDynamicNumberingIdState,
    reserveNextNumberingIdPair,
    mergeNumberingXmlBySchemaOrder,
    remapNumberingPayloadForDocument,
    overwriteParagraphNumIds,
    extractFirstParagraphNumId,
    buildExplicitDecimalMultilevelNumberingXml,
    inferTableReplacementParagraphBlock,
    resolveParagraphRangeByRefs
} from '../src/taskpane/modules/reconciliation/standalone.js';

const DEMO_VERSION = '2026-02-15-chat-docx-preview-19';
const GEMINI_API_KEY_STORAGE_KEY = 'browserDemo.geminiApiKey';
const DEMO_MARKERS = [
    'DEMO_TEXT_TARGET',
    'DEMO FORMAT TARGET',
    'DEMO_LIST_TARGET',
    'DEMO_TABLE_TARGET'
];
const ALLOWED_HIGHLIGHT_COLORS = ['yellow', 'green', 'cyan', 'magenta', 'blue', 'red'];
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NUMBERING_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const NUMBERING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';
const COMMENTS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const COMMENTS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';

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

// ── State ──────────────────────────────────────────────
let currentZip = null;           // JSZip instance of the working document
let documentParagraphs = [];     // [{ index, text }] extracted from current docx
let chatHistory = [];            // Gemini multi-turn history [{ role, parts }]
let operationCount = 0;          // total operations applied across turns
let previewRenderer = null;      // docxjs renderAsync function
let previewRenderToken = 0;      // guards stale async preview writes

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

// ── XML Helpers (unchanged from original demo) ─────────
function parseXmlStrict(xmlText, label) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) throw new Error(`[XML parse error] ${label}: ${parseError.textContent || 'Unknown'}`);
    return xmlDoc;
}

function isSectionPropertiesElement(node) {
    return !!node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'sectPr';
}
function getBodyElement(xmlDoc) {
    return xmlDoc.getElementsByTagNameNS('*', 'body')[0] || null;
}
function getDirectSectionProperties(body) {
    for (const child of Array.from(body.childNodes)) {
        if (isSectionPropertiesElement(child)) return child;
    }
    return null;
}
function insertBodyElementBeforeSectPr(body, element) {
    const sectPr = getDirectSectionProperties(body);
    if (sectPr) body.insertBefore(element, sectPr);
    else body.appendChild(element);
}
function normalizeBodySectionOrder(xmlDoc) {
    const body = getBodyElement(xmlDoc);
    if (!body) return;
    const sectPr = getDirectSectionProperties(body);
    if (!sectPr) return;
    let cursor = sectPr.nextSibling;
    while (cursor) {
        const next = cursor.nextSibling;
        if (cursor.nodeType === 1) body.insertBefore(cursor, sectPr);
        cursor = next;
    }
}

// ── Sanitize nested paragraphs ─────────────────────────
/**
 * After redlining table content, the engine can produce nested w:p elements
 * inside table cells (w:tc > w:p > w:p). This flattens them by promoting
 * the inner w:p's children into the outer w:p, then removing the inner w:p.
 */
function sanitizeNestedParagraphs(xmlDoc) {
    const tcs = xmlDoc.getElementsByTagNameNS(NS_W, 'tc');
    let fixed = 0;
    for (const tc of Array.from(tcs)) {
        const outerParagraphs = Array.from(tc.childNodes).filter(
            n => n.nodeType === 1 && n.namespaceURI === NS_W && n.localName === 'p'
        );
        for (const outerP of outerParagraphs) {
            const innerParagraphs = Array.from(outerP.childNodes).filter(
                n => n.nodeType === 1 && n.namespaceURI === NS_W && n.localName === 'p'
            );
            for (const innerP of innerParagraphs) {
                // Move all children of the inner <w:p> into the parent <w:tc>, before the outer <w:p>
                // Then remove the inner <w:p> from the outer <w:p>
                // Strategy: promote innerP to be a sibling of outerP in the tc
                tc.insertBefore(innerP, outerP);
                fixed++;
            }
        }
    }
    if (fixed > 0) log(`[Sanitize] Fixed ${fixed} nested w:p element(s) in table cells`);
}

// ── Paragraph helpers ──────────────────────────────────
function getParagraphText(paragraph) {
    return getParagraphTextFromOxml(paragraph);
}

function findParagraphByStrictText(xmlDoc, targetText) {
    return findParagraphByStrictTextShared(xmlDoc, targetText);
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

function resolveTargetParagraph(xmlDoc, targetText, targetRef, opType, runtimeContext = null) {
    return resolveTargetParagraphWithSnapshotShared(xmlDoc, {
        targetText,
        targetRef,
        opType,
        targetRefSnapshot: runtimeContext?.targetRefSnapshot || null,
        onInfo: message => log(message),
        onWarn: message => log(message)
    });
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

// ── Package extraction helpers ─────────────────────────
function getPartName(partElement) {
    return partElement.getAttribute('pkg:name') || partElement.getAttribute('name') || '';
}

function extractFromPackage(packageXml) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const pkgDoc = parser.parseFromString(packageXml, 'application/xml');
    const parts = Array.from(pkgDoc.getElementsByTagNameNS('*', 'part'));
    const documentPart = parts.find(p => getPartName(p) === '/word/document.xml');
    if (!documentPart) throw new Error('Package output missing /word/document.xml part');
    const xmlData = documentPart.getElementsByTagNameNS('*', 'xmlData')[0];
    if (!xmlData) throw new Error('Package document part missing pkg:xmlData');
    const documentNode = Array.from(xmlData.childNodes).find(n => n.nodeType === 1);
    if (!documentNode) throw new Error('Package document part missing XML payload');
    const body = documentNode.getElementsByTagNameNS('*', 'body')[0];
    const replacementNodes = body
        ? Array.from(body.childNodes).filter(n => n.nodeType === 1 && !isSectionPropertiesElement(n))
        : [documentNode];
    const numberingPart = parts.find(p => getPartName(p) === '/word/numbering.xml');
    let numberingXml = null;
    if (numberingPart) {
        const nd = numberingPart.getElementsByTagNameNS('*', 'xmlData')[0];
        const nn = nd ? Array.from(nd.childNodes).find(n => n.nodeType === 1) : null;
        if (nn) numberingXml = serializer.serializeToString(nn);
    }
    return { replacementNodes, numberingXml };
}

function extractReplacementNodes(outputOxml) {
    if (typeof outputOxml !== 'string' || !outputOxml.trim()) {
        throw new Error('Reconciliation engine returned no OOXML payload for this operation');
    }
    const parser = new DOMParser();
    if (outputOxml.includes('<pkg:package')) return extractFromPackage(outputOxml);
    if (outputOxml.includes('<w:document')) {
        const doc = parser.parseFromString(outputOxml, 'application/xml');
        const body = doc.getElementsByTagNameNS('*', 'body')[0];
        const nodes = body
            ? Array.from(body.childNodes).filter(n => n.nodeType === 1 && !isSectionPropertiesElement(n))
            : Array.from(doc.childNodes).filter(n => n.nodeType === 1);
        return { replacementNodes: nodes, numberingXml: null };
    }
    const wrapped = `<root xmlns:w="${NS_W}">${outputOxml}</root>`;
    const fragmentDoc = parser.parseFromString(wrapped, 'application/xml');
    const nodes = Array.from(fragmentDoc.documentElement.childNodes).filter(n => n.nodeType === 1);
    return { replacementNodes: nodes, numberingXml: null };
}

function getDirectWordChild(element, localName) {
    if (!element) return null;
    return Array.from(element.childNodes || []).find(
        node => node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === localName
    ) || null;
}

function getNextTrackedChangeId(xmlDoc) {
    let maxId = 999;
    const revisionNodes = [
        ...Array.from(xmlDoc.getElementsByTagNameNS(NS_W, 'ins')),
        ...Array.from(xmlDoc.getElementsByTagNameNS(NS_W, 'del'))
    ];
    for (const node of revisionNodes) {
        const raw = node.getAttribute('w:id') || node.getAttribute('id') || '';
        const parsed = Number.parseInt(raw, 10);
        if (Number.isFinite(parsed)) maxId = Math.max(maxId, parsed);
    }
    return maxId + 1;
}

function ensureListProperties(xmlDoc, paragraph, ilvl, numId) {
    let pPr = getDirectWordChild(paragraph, 'pPr');
    if (!pPr) {
        pPr = xmlDoc.createElementNS(NS_W, 'w:pPr');
        paragraph.insertBefore(pPr, paragraph.firstChild);
    }

    let numPr = getDirectWordChild(pPr, 'numPr');
    if (!numPr) {
        numPr = xmlDoc.createElementNS(NS_W, 'w:numPr');
        pPr.appendChild(numPr);
    }

    let ilvlEl = getDirectWordChild(numPr, 'ilvl');
    if (!ilvlEl) {
        ilvlEl = xmlDoc.createElementNS(NS_W, 'w:ilvl');
        numPr.appendChild(ilvlEl);
    }
    ilvlEl.setAttribute('w:val', String(Math.max(0, Number.parseInt(ilvl, 10) || 0)));

    let numIdEl = getDirectWordChild(numPr, 'numId');
    if (!numIdEl) {
        numIdEl = xmlDoc.createElementNS(NS_W, 'w:numId');
        numPr.appendChild(numIdEl);
    }
    numIdEl.setAttribute('w:val', String(numId));
}

function buildInsertedListParagraph(xmlDoc, anchorParagraph, entry, revisionId, author, dateIso) {
    const paragraph = xmlDoc.createElementNS(NS_W, 'w:p');

    const anchorPPr = getDirectWordChild(anchorParagraph, 'pPr');
    if (anchorPPr) {
        paragraph.appendChild(anchorPPr.cloneNode(true));
    }
    ensureListProperties(xmlDoc, paragraph, entry.ilvl, entry.numId);

    const ins = xmlDoc.createElementNS(NS_W, 'w:ins');
    ins.setAttribute('w:id', String(revisionId));
    ins.setAttribute('w:author', author || 'Browser Demo AI');
    ins.setAttribute('w:date', dateIso);

    const run = xmlDoc.createElementNS(NS_W, 'w:r');
    const anchorFirstRun = Array.from(anchorParagraph.getElementsByTagNameNS(NS_W, 'r'))[0] || null;
    const anchorRunPr = anchorFirstRun ? getDirectWordChild(anchorFirstRun, 'rPr') : null;
    if (anchorRunPr) {
        run.appendChild(anchorRunPr.cloneNode(true));
    }

    const textNode = xmlDoc.createElementNS(NS_W, 'w:t');
    const safeText = String(entry.text || '').trim();
    if (/^\s|\s$/.test(safeText)) textNode.setAttribute('xml:space', 'preserve');
    textNode.textContent = safeText;
    run.appendChild(textNode);
    ins.appendChild(run);
    paragraph.appendChild(ins);

    return paragraph;
}

function serializeParagraphRangeAsDocument(paragraphs, serializer) {
    const paragraphXml = (paragraphs || [])
        .map(p => serializer.serializeToString(p))
        .join('');
    return `<w:document xmlns:w="${NS_W}"><w:body>${paragraphXml}</w:body></w:document>`;
}

async function tryExplicitDecimalHeaderListConversion({
    xmlDoc,
    serializer,
    targetParagraph,
    currentParagraphText,
    modifiedText,
    author,
    runtimeContext
}) {
    if (!targetParagraph) return null;
    const scopedParagraphOxml = serializer.serializeToString(targetParagraph);
    const explicitPlan = buildSingleLineListStructuralFallbackPlan({
        oxml: scopedParagraphOxml,
        originalText: currentParagraphText,
        modifiedText,
        allowExistingList: false
    });
    if (
        !explicitPlan ||
        explicitPlan.numberingKey !== 'numbered:decimal:single' ||
        !Number.isInteger(explicitPlan.startAt) ||
        explicitPlan.startAt < 1
    ) {
        return null;
    }

    const strippedContent = stripSingleLineListMarkerPrefix(explicitPlan.listInput || modifiedText);
    if (!strippedContent) return null;

    log('[List] Applying explicit numeric header conversion with direct list binding.');
    const redlineResult = await applyRedlineToOxml(
        serializer.serializeToString(targetParagraph),
        currentParagraphText,
        strippedContent,
        {
            author,
            generateRedlines: true
        }
    );
    if (!redlineResult?.hasChanges || typeof redlineResult?.oxml !== 'string') return null;

    const extracted = extractReplacementNodes(redlineResult.oxml);
    const replacementNodes = extracted.replacementNodes;
    const numberingAction = resolveSingleLineListFallbackNumberingAction(
        explicitPlan,
        runtimeContext?.listFallbackSequenceState || null
    );

    const explicitStart = explicitPlan.startAt;
    const numberingState = runtimeContext?.numberingIdState || null;
    let appliedNumId = null;
    let numberingXml = null;

    if (numberingAction.type === 'explicitReuse' && numberingAction.numId) {
        appliedNumId = String(numberingAction.numId);
        log(`[List] Reusing explicit-start list sequence (${numberingAction.numberingKey} -> numId ${appliedNumId}, next ${explicitStart + 1}).`);
    } else {
        const reservedPair = reserveNextNumberingIdPair(numberingState);
        if (!reservedPair) return null;

        appliedNumId = String(reservedPair.numId);
        numberingXml = buildExplicitDecimalMultilevelNumberingXml(
            reservedPair.numId,
            reservedPair.abstractNumId,
            explicitStart
        );

        if (numberingAction.type === 'explicitStartNew') {
            log(`[List] Started explicit-start list sequence (${numberingAction.numberingKey} -> numId ${appliedNumId}).`);
        }
        log(`[List] Using isolated explicit-start numbering (start ${explicitStart}, numId ${appliedNumId}, abstractNumId ${reservedPair.abstractNumId}).`);
    }

    if (explicitPlan.numberingKey && runtimeContext?.listFallbackSharedNumIdByKey instanceof Map) {
        runtimeContext.listFallbackSharedNumIdByKey.delete(explicitPlan.numberingKey);
    }

    if (numberingAction.type === 'explicitStartNew' || numberingAction.type === 'explicitReuse') {
        recordSingleLineListFallbackExplicitSequence(
            runtimeContext?.listFallbackSequenceState || null,
            numberingAction.numberingKey || explicitPlan.numberingKey,
            appliedNumId,
            explicitStart
        );
    } else {
        clearSingleLineListFallbackExplicitSequence(
            runtimeContext?.listFallbackSequenceState || null,
            numberingAction.numberingKey || explicitPlan.numberingKey
        );
    }

    enforceListBindingOnParagraphNodes(replacementNodes, {
        numId: appliedNumId,
        ilvl: 0,
        clearParagraphPropertyChanges: true,
        removeListPropertyNode: true
    });

    const parent = targetParagraph.parentNode;
    if (!parent) return null;
    for (const node of replacementNodes) parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml };
}

async function trySingleParagraphListStructuralFallback({
    xmlDoc,
    serializer,
    targetParagraph,
    currentParagraphText,
    modifiedText,
    author,
    runtimeContext
}) {
    if (!targetParagraph) return null;

    const scopedParagraphOxml = serializer.serializeToString(targetParagraph);
    const fallbackPlan = buildSingleLineListStructuralFallbackPlan({
        oxml: scopedParagraphOxml,
        originalText: currentParagraphText,
        modifiedText,
        allowExistingList: false
    });
    if (!fallbackPlan) return null;

    log('[List] No textual diff but list marker detected; forcing structural list conversion fallback.');
    const fallbackResult = await executeSingleLineListStructuralFallback(fallbackPlan, {
        author,
        generateRedlines: true,
        setAbstractStartOverride: false
    });
    if (!fallbackResult?.hasChanges || !fallbackResult?.oxml) {
        log('[List] Structural list fallback produced no valid OOXML payload.');
        return null;
    }

    const extracted = extractReplacementNodes(fallbackResult.oxml);
    let replacementNodes = extracted.replacementNodes;
    let numberingXml = extracted.numberingXml || fallbackResult?.numberingXml || null;
    const hasExplicitStartAt = Number.isInteger(fallbackPlan?.startAt) && fallbackPlan.startAt > 0;
    const numberingKey = fallbackResult?.listStructuralFallbackKey || fallbackPlan?.numberingKey || null;
    const numberingAction = resolveSingleLineListFallbackNumberingAction(
        fallbackPlan,
        runtimeContext?.listFallbackSequenceState || null
    );
    if (hasExplicitStartAt) {
        const explicitStart = fallbackPlan.startAt;
        let explicitNumIdForBinding = null;
        const numberingState = runtimeContext?.numberingIdState || null;
        if (numberingAction.type === 'explicitReuse' && numberingAction.numId) {
            explicitNumIdForBinding = String(numberingAction.numId);
            numberingXml = null;
            log(`[List] Reusing explicit-start list sequence (${numberingAction.numberingKey} -> numId ${explicitNumIdForBinding}, next ${explicitStart + 1}).`);
        } else if (numberingState) {
            const reservedPair = reserveNextNumberingIdPair(numberingState);
            if (!reservedPair) return null;
            overwriteParagraphNumIds(replacementNodes, reservedPair.numId);
            explicitNumIdForBinding = String(reservedPair.numId);
            numberingXml = buildExplicitDecimalMultilevelNumberingXml(
                reservedPair.numId,
                reservedPair.abstractNumId,
                explicitStart
            );
            if (numberingAction.type === 'explicitStartNew') {
                log(`[List] Started explicit-start list sequence (${numberingAction.numberingKey} -> numId ${reservedPair.numId}).`);
            }
            log(`[List] Using isolated explicit-start numbering (start ${explicitStart}, numId ${reservedPair.numId}, abstractNumId ${reservedPair.abstractNumId}).`);
        } else {
            const generatedNumId = extractFirstParagraphNumId(replacementNodes);
            explicitNumIdForBinding = generatedNumId ? String(generatedNumId) : null;
            log(`[List] Using isolated list numbering with explicit start ${explicitStart}${generatedNumId ? ` (numId ${generatedNumId})` : ''}.`);
        }

        if (numberingKey && runtimeContext?.listFallbackSharedNumIdByKey instanceof Map) {
            runtimeContext.listFallbackSharedNumIdByKey.delete(numberingKey);
        }
        if (numberingAction.type === 'explicitStartNew' || numberingAction.type === 'explicitReuse') {
            recordSingleLineListFallbackExplicitSequence(
                runtimeContext?.listFallbackSequenceState || null,
                numberingAction.numberingKey || numberingKey,
                explicitNumIdForBinding,
                explicitStart
            );
        } else {
            clearSingleLineListFallbackExplicitSequence(
                runtimeContext?.listFallbackSequenceState || null,
                numberingAction.numberingKey || numberingKey
            );
        }

        if (explicitNumIdForBinding) {
            enforceListBindingOnParagraphNodes(replacementNodes, {
                numId: explicitNumIdForBinding,
                ilvl: 0,
                clearParagraphPropertyChanges: true,
                removeListPropertyNode: true
            });
        }
    } else {
        if (numberingXml && runtimeContext?.numberingIdState) {
            const normalizedNumbering = remapNumberingPayloadForDocument(numberingXml, replacementNodes, runtimeContext.numberingIdState);
            replacementNodes = normalizedNumbering.replacementNodes;
            numberingXml = normalizedNumbering.numberingXml;
        }
        clearSingleLineListFallbackExplicitSequence(
            runtimeContext?.listFallbackSequenceState || null,
            numberingAction.numberingKey || numberingKey
        );
    }

    if (!hasExplicitStartAt && runtimeContext?.listFallbackSharedNumIdByKey instanceof Map) {
        const sharedNumId = numberingKey ? runtimeContext.listFallbackSharedNumIdByKey.get(numberingKey) : null;
        if (sharedNumId) {
            overwriteParagraphNumIds(replacementNodes, sharedNumId);
            numberingXml = null;
            log(`[List] Reusing shared list numbering (${numberingKey} -> numId ${sharedNumId}).`);
        } else if (numberingKey) {
            const generatedNumId = extractFirstParagraphNumId(replacementNodes);
            if (generatedNumId) {
                runtimeContext.listFallbackSharedNumIdByKey.set(numberingKey, generatedNumId);
                log(`[List] Captured shared list numbering (${numberingKey} -> numId ${generatedNumId}).`);
            }
        }
    }

    const parent = targetParagraph.parentNode;
    if (!parent) return null;
    for (const node of replacementNodes) parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml };
}

// ── Apply operations (per-paragraph) ───────────────────
async function applyToParagraphByExactText(documentXml, targetText, modifiedText, author, targetRef = null, targetEndRef = null, runtimeContext = null) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'redline', runtimeContext);
    const targetParagraph = resolved.paragraph;
    const currentParagraphText = getParagraphText(targetParagraph).trim();
    const containingTable = findContainingWordElement(targetParagraph, 'tbl');
    const synthesizedTableMarkdown = containingTable
        ? synthesizeTableMarkdownFromMultilineCellEdit(targetParagraph, modifiedText, {
            tableElement: containingTable,
            currentParagraphText,
            onInfo: message => log(message),
            onWarn: message => log(message)
        })
        : null;
    let effectiveModifiedText = synthesizedTableMarkdown || modifiedText;
    const useTableScope = !!containingTable && isMarkdownTableText(effectiveModifiedText);
    const isTableMarkdownEdit = isMarkdownTableText(effectiveModifiedText);
    const explicitRangeParagraphs = targetEndRef
        ? resolveParagraphRangeByRefs(xmlDoc, targetRef, targetEndRef, {
            opType: 'redline',
            targetRefSnapshot: runtimeContext?.targetRefSnapshot || null,
            onInfo: message => log(message),
            onWarn: message => log(message)
        })
        : null;
    let inferredTableRangeParagraphs = null;
    if (!explicitRangeParagraphs && !useTableScope && isTableMarkdownEdit) {
        inferredTableRangeParagraphs = inferTableReplacementParagraphBlock(targetParagraph, {
            getParagraphText
        });
        if (inferredTableRangeParagraphs?.length > 1) {
            log(`[Table] Heuristic range expansion selected ${inferredTableRangeParagraphs.length} paragraph(s) for replacement.`);
        }
    }
    const targetListInfo = getParagraphListInfo(targetParagraph);
    if (
        targetListInfo &&
        typeof effectiveModifiedText === 'string' &&
        !effectiveModifiedText.includes('\n')
    ) {
        const strippedListPrefix = stripRedundantLeadingListMarkers(effectiveModifiedText);
        if (strippedListPrefix && strippedListPrefix !== effectiveModifiedText.trim()) {
            log('[List] Stripped redundant manual list marker prefix from single-line list item edit.');
            effectiveModifiedText = strippedListPrefix;
        }
    }
    if (useTableScope) {
        log('[Table] Markdown table edit detected in table cell target; applying reconciliation at table scope.');
    }

    const insertionOnlyPlan = !useTableScope
        ? planListInsertionOnlyEdit(targetParagraph, effectiveModifiedText, {
            currentParagraphText,
            onInfo: message => log(message),
            onWarn: message => log(message)
        })
        : null;
    if (insertionOnlyPlan && insertionOnlyPlan.entries.length > 0) {
        log(`[List] Applying insertion-only list redline heuristic (${insertionOnlyPlan.entries.length} new item(s)).`);
        for (const entry of insertionOnlyPlan.entries) {
            log(`[List] Insertion entry resolved: ilvl=${entry.ilvl}, markerType=${entry.markerType}, text="${String(entry.text || '').slice(0, 80)}${String(entry.text || '').length > 80 ? '…' : ''}"`);
        }
        const parent = targetParagraph.parentNode;
        if (!parent) throw new Error('Target paragraph has no parent for list insertion');
        const insertionPoint = targetParagraph.nextSibling;
        const dateIso = new Date().toISOString();
        let revisionId = getNextTrackedChangeId(xmlDoc);
        for (const entry of insertionOnlyPlan.entries) {
            const listParagraph = buildInsertedListParagraph(
                xmlDoc,
                targetParagraph,
                { ...entry, numId: insertionOnlyPlan.numId },
                revisionId++,
                author,
                dateIso
            );
            parent.insertBefore(listParagraph, insertionPoint);
        }
        normalizeBodySectionOrder(xmlDoc);
        return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml: null };
    }

    const listScopeEdit = !useTableScope
        ? synthesizeExpandedListScopeEdit(targetParagraph, effectiveModifiedText, {
            currentParagraphText,
            onInfo: message => log(message),
            onWarn: message => log(message)
        })
        : null;
    const useListScope = !!listScopeEdit && Array.isArray(listScopeEdit.paragraphs) && listScopeEdit.paragraphs.length > 0;
    if (useListScope) {
        effectiveModifiedText = listScopeEdit.modifiedText;
    }

    if (!useTableScope && !useListScope) {
        const explicitHeaderListConversion = await tryExplicitDecimalHeaderListConversion({
            xmlDoc,
            serializer,
            targetParagraph,
            currentParagraphText,
            modifiedText: effectiveModifiedText,
            author,
            runtimeContext
        });
        if (explicitHeaderListConversion) return explicitHeaderListConversion;

        const listFallback = await trySingleParagraphListStructuralFallback({
            xmlDoc,
            serializer,
            targetParagraph,
            currentParagraphText,
            modifiedText: effectiveModifiedText,
            author,
            runtimeContext
        });
        if (listFallback) return listFallback;
    }

    const originalTextForApply = useListScope
        ? listScopeEdit.originalText
        : (
            explicitRangeParagraphs
                ? explicitRangeParagraphs.map(p => getParagraphText(p)).join('\n')
                : (inferredTableRangeParagraphs
                    ? inferredTableRangeParagraphs.map(p => getParagraphText(p)).join('\n')
                    : (currentParagraphText || targetText))
        );
    const scopedXml = useTableScope
        ? serializer.serializeToString(containingTable)
        : (
            useListScope
                ? serializeParagraphRangeAsDocument(listScopeEdit.paragraphs, serializer)
                : (
                    explicitRangeParagraphs
                        ? serializeParagraphRangeAsDocument(explicitRangeParagraphs, serializer)
                        : (inferredTableRangeParagraphs
                            ? serializeParagraphRangeAsDocument(inferredTableRangeParagraphs, serializer)
                            : serializer.serializeToString(targetParagraph))
                )
        );

    const result = isTableMarkdownEdit
        ? await reconcileMarkdownTableOoxml(scopedXml, originalTextForApply, effectiveModifiedText, {
            author,
            generateRedlines: true,
            // Preserve table wrapper when table markdown should reconcile the full table.
            _isolatedTableCell: useTableScope
        })
        : await applyRedlineToOxml(scopedXml, originalTextForApply, effectiveModifiedText, {
            author,
            generateRedlines: true,
            // Preserve table wrapper when table markdown should reconcile the full table.
            _isolatedTableCell: useTableScope
        });
    if (!result?.hasChanges) return { documentXml, hasChanges: false, numberingXml: null };
    if (result.useNativeApi && !result.oxml) {
        const warning = 'Format-only fallback requires native Word API; browser demo skipped this operation.';
        log(`[WARN] ${warning}`);
        return { documentXml, hasChanges: false, numberingXml: null, warnings: [warning] };
    }
    if (typeof result.oxml !== 'string') {
        throw new Error('Reconciliation engine did not return OOXML for a changed redline operation');
    }
    const extracted = extractReplacementNodes(result.oxml);
    let replacementNodes = extracted.replacementNodes;
    let numberingXml = extracted.numberingXml;
    if (numberingXml && runtimeContext?.numberingIdState) {
        const normalizedNumbering = remapNumberingPayloadForDocument(numberingXml, replacementNodes, runtimeContext.numberingIdState);
        replacementNodes = normalizedNumbering.replacementNodes;
        numberingXml = normalizedNumbering.numberingXml;
    }
    const scopeNodes = useTableScope
        ? [containingTable]
        : (
            useListScope
                ? listScopeEdit.paragraphs
                : (
                    explicitRangeParagraphs
                        ? explicitRangeParagraphs
                        : (inferredTableRangeParagraphs || [targetParagraph])
                )
        );
    const anchorNode = scopeNodes[0];
    const parent = anchorNode.parentNode;
    for (const node of replacementNodes) parent.insertBefore(xmlDoc.importNode(node, true), anchorNode);
    for (const scopeNode of scopeNodes) {
        if (scopeNode && scopeNode.parentNode === parent) parent.removeChild(scopeNode);
    }
    normalizeBodySectionOrder(xmlDoc);
    return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml };
}

async function applyHighlightToParagraphByExactText(documentXml, targetText, textToHighlight, color, author, targetRef = null, runtimeContext = null) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'highlight', runtimeContext);
    const targetParagraph = resolved.paragraph;
    const paragraphXml = serializer.serializeToString(targetParagraph);
    const highlightedXml = applyHighlightToOoxml(paragraphXml, textToHighlight, color, { generateRedlines: true, author });
    if (!highlightedXml || highlightedXml === paragraphXml) return { documentXml, hasChanges: false };
    const { replacementNodes } = extractReplacementNodes(highlightedXml);
    const parent = targetParagraph.parentNode;
    for (const node of replacementNodes) parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true };
}

async function applyCommentToParagraphByExactText(documentXml, targetText, textToComment, commentContent, author, targetRef = null, runtimeContext = null) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'comment', runtimeContext);
    const targetParagraph = resolved.paragraph;
    const paragraphXml = serializer.serializeToString(targetParagraph);
    const commentResult = injectCommentsIntoOoxml(paragraphXml, [{ paragraphIndex: 1, textToFind: textToComment, commentContent }], { author });
    if (!commentResult.commentsApplied) return { documentXml, hasChanges: false, commentsXml: null, warnings: commentResult.warnings || [] };
    const { replacementNodes } = extractReplacementNodes(commentResult.oxml);
    const parent = targetParagraph.parentNode;
    for (const node of replacementNodes) parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, commentsXml: commentResult.commentsXml || null, warnings: commentResult.warnings || [] };
}

async function runOperation(documentXml, op, author, runtimeContext = null) {
    if (op.type === 'highlight') return applyHighlightToParagraphByExactText(documentXml, op.target, op.textToHighlight, op.color, author, op.targetRef, runtimeContext);
    if (op.type === 'comment') return applyCommentToParagraphByExactText(documentXml, op.target, op.textToComment, op.commentContent, author, op.targetRef, runtimeContext);
    return applyToParagraphByExactText(documentXml, op.target, op.modified, author, op.targetRef, op.targetEndRef, runtimeContext);
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
    const incomingPayloads = (Array.isArray(numberingXmlList) ? numberingXmlList : [numberingXmlList]).filter(Boolean);
    if (incomingPayloads.length === 0) return;

    const existing = await zip.file('word/numbering.xml')?.async('string');
    let mergedNumberingXml = existing || null;
    for (const incomingNumbering of incomingPayloads) {
        mergedNumberingXml = mergedNumberingXml
            ? mergeNumberingXml(mergedNumberingXml, incomingNumbering)
            : incomingNumbering;
    }

    if (!existing) log('[Demo] Adding numbering.xml');
    else log('[Demo] Merging numbering.xml payload(s) into existing numbering definitions');
    zip.file('word/numbering.xml', mergedNumberingXml);

    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const ctText = await zip.file('[Content_Types].xml')?.async('string');
    if (ctText) {
        const ctDoc = parser.parseFromString(ctText, 'application/xml');
        const overrides = Array.from(ctDoc.getElementsByTagNameNS('*', 'Override'));
        if (!overrides.some(o => (o.getAttribute('PartName') || '').toLowerCase() === '/word/numbering.xml')) {
            const ov = ctDoc.createElementNS(NS_CT, 'Override');
            ov.setAttribute('PartName', '/word/numbering.xml');
            ov.setAttribute('ContentType', NUMBERING_CONTENT_TYPE);
            ctDoc.documentElement.appendChild(ov);
            zip.file('[Content_Types].xml', serializer.serializeToString(ctDoc));
        }
    }
    const relsPath = 'word/_rels/document.xml.rels';
    const relsText = await zip.file(relsPath)?.async('string');
    if (relsText) {
        const relsDoc = parser.parseFromString(relsText, 'application/xml');
        const relsRoot = relsDoc.getElementsByTagNameNS('*', 'Relationships')[0] || relsDoc.documentElement;
        const rels = Array.from(relsRoot.getElementsByTagNameNS('*', 'Relationship'));
        if (!rels.some(r => (r.getAttribute('Type') || '') === NUMBERING_REL_TYPE)) {
            let max = 0;
            for (const rel of rels) { const n = parseInt((rel.getAttribute('Id') || '').replace(/^rId/i, ''), 10); if (!Number.isNaN(n)) max = Math.max(max, n); }
            const rel = relsDoc.createElementNS(NS_RELS, 'Relationship');
            rel.setAttribute('Id', `rId${max + 1}`);
            rel.setAttribute('Type', NUMBERING_REL_TYPE);
            rel.setAttribute('Target', 'numbering.xml');
            relsRoot.appendChild(rel);
            zip.file(relsPath, serializer.serializeToString(relsDoc));
        }
    }
}

async function ensureCommentsArtifacts(zip, commentsXml) {
    if (!commentsXml) return;
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const commentsPath = 'word/comments.xml';
    const existingText = await zip.file(commentsPath)?.async('string');
    if (!existingText) { log('[Demo] Adding comments.xml'); zip.file(commentsPath, commentsXml); }
    else {
        const existingDoc = parseXmlStrict(existingText, 'word/comments.xml (existing)');
        const incomingDoc = parseXmlStrict(commentsXml, 'word/comments.xml (incoming)');
        const existingRoot = existingDoc.documentElement;
        const existingIds = new Set(Array.from(existingRoot.getElementsByTagNameNS(NS_W, 'comment')).map(c => c.getAttribute('w:id') || c.getAttribute('id')).filter(Boolean));
        for (const ic of Array.from(incomingDoc.documentElement.getElementsByTagNameNS(NS_W, 'comment'))) {
            const id = ic.getAttribute('w:id') || ic.getAttribute('id');
            if (id && existingIds.has(id)) throw new Error(`Duplicate comment id: ${id}`);
            existingRoot.appendChild(existingDoc.importNode(ic, true));
        }
        zip.file(commentsPath, serializer.serializeToString(existingDoc));
    }
    const ctText = await zip.file('[Content_Types].xml')?.async('string');
    if (ctText) {
        const ctDoc = parser.parseFromString(ctText, 'application/xml');
        if (!Array.from(ctDoc.getElementsByTagNameNS('*', 'Override')).some(o => (o.getAttribute('PartName') || '').toLowerCase() === '/word/comments.xml')) {
            const ov = ctDoc.createElementNS(NS_CT, 'Override');
            ov.setAttribute('PartName', '/word/comments.xml');
            ov.setAttribute('ContentType', COMMENTS_CONTENT_TYPE);
            ctDoc.documentElement.appendChild(ov);
            zip.file('[Content_Types].xml', serializer.serializeToString(ctDoc));
        }
    }
    const relsPath = 'word/_rels/document.xml.rels';
    const relsText = await zip.file(relsPath)?.async('string');
    if (relsText) {
        const relsDoc = parser.parseFromString(relsText, 'application/xml');
        const relsRoot = relsDoc.getElementsByTagNameNS('*', 'Relationships')[0] || relsDoc.documentElement;
        const rels = Array.from(relsRoot.getElementsByTagNameNS('*', 'Relationship'));
        if (!rels.some(r => (r.getAttribute('Type') || '') === COMMENTS_REL_TYPE)) {
            let max = 0;
            for (const rel of rels) { const n = parseInt((rel.getAttribute('Id') || '').replace(/^rId/i, ''), 10); if (!Number.isNaN(n)) max = Math.max(max, n); }
            const rel = relsDoc.createElementNS(NS_RELS, 'Relationship');
            rel.setAttribute('Id', `rId${max + 1}`);
            rel.setAttribute('Type', COMMENTS_REL_TYPE);
            rel.setAttribute('Target', 'comments.xml');
            relsRoot.appendChild(rel);
            zip.file(relsPath, serializer.serializeToString(relsDoc));
        }
    }
}

// ── Validation ─────────────────────────────────────────
async function validateOutputDocx(zip) {
    const documentXml = await zip.file('word/document.xml')?.async('string');
    if (!documentXml) throw new Error('Validation failed: missing word/document.xml');
    const documentDoc = parseXmlStrict(documentXml, 'word/document.xml');
    normalizeBodySectionOrder(documentDoc);
    const body = getBodyElement(documentDoc);
    if (!body) throw new Error('Validation failed: word/document.xml has no w:body');

    const directBodyElements = Array.from(body.childNodes).filter(n => n.nodeType === 1);
    const sectPrIndexes = directBodyElements.map((node, idx) => ({ node, idx })).filter(({ node }) => isSectionPropertiesElement(node)).map(({ idx }) => idx);
    if (sectPrIndexes.length > 1) throw new Error('Validation failed: multiple body-level w:sectPr');
    if (sectPrIndexes.length === 1 && sectPrIndexes[0] !== directBodyElements.length - 1) throw new Error('Validation failed: w:sectPr not last');

    const tcs = documentDoc.getElementsByTagNameNS(NS_W, 'tc');
    for (const tc of Array.from(tcs)) {
        for (const child of Array.from(tc.childNodes).filter(n => n.nodeType === 1)) {
            if (child.namespaceURI === NS_W && child.localName === 'p') {
                if (Array.from(child.childNodes).find(n => n.nodeType === 1 && n.namespaceURI === NS_W && n.localName === 'p')) throw new Error('Validation failed: nested w:p');
            }
        }
    }

    const hasNumberingUsage = documentDoc.getElementsByTagNameNS(NS_W, 'numPr').length > 0;
    const hasCommentUsage = documentDoc.getElementsByTagNameNS(NS_W, 'commentRangeStart').length > 0 || documentDoc.getElementsByTagNameNS(NS_W, 'commentRangeEnd').length > 0 || documentDoc.getElementsByTagNameNS(NS_W, 'commentReference').length > 0;
    const numberingXml = await zip.file('word/numbering.xml')?.async('string');
    const commentsXml = await zip.file('word/comments.xml')?.async('string');
    if (numberingXml) parseXmlStrict(numberingXml, 'word/numbering.xml');
    else if (hasNumberingUsage) throw new Error('Validation failed: numbering used but part missing');
    if (commentsXml) parseXmlStrict(commentsXml, 'word/comments.xml');
    else if (hasCommentUsage) throw new Error('Validation failed: comments used but part missing');

    const ctXml = await zip.file('[Content_Types].xml')?.async('string');
    if (!ctXml) throw new Error('Validation failed: missing [Content_Types].xml');
    const ctDoc = parseXmlStrict(ctXml, '[Content_Types].xml');
    const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string');
    if (!relsXml) throw new Error('Validation failed: missing document.xml.rels');
    const relsDoc = parseXmlStrict(relsXml, 'document.xml.rels');

    if (numberingXml) {
        if (!Array.from(ctDoc.getElementsByTagNameNS('*', 'Override')).some(o => (o.getAttribute('PartName') || '').toLowerCase() === '/word/numbering.xml' && (o.getAttribute('ContentType') || '') === NUMBERING_CONTENT_TYPE)) throw new Error('Validation failed: numbering CT override missing');
        if (!Array.from(relsDoc.getElementsByTagNameNS('*', 'Relationship')).some(r => (r.getAttribute('Type') || '') === NUMBERING_REL_TYPE)) throw new Error('Validation failed: numbering rel missing');
    }
    if (commentsXml) {
        if (!Array.from(ctDoc.getElementsByTagNameNS('*', 'Override')).some(o => (o.getAttribute('PartName') || '').toLowerCase() === '/word/comments.xml' && (o.getAttribute('ContentType') || '') === COMMENTS_CONTENT_TYPE)) throw new Error('Validation failed: comments CT override missing');
        if (!Array.from(relsDoc.getElementsByTagNameNS('*', 'Relationship')).some(r => (r.getAttribute('Type') || '') === COMMENTS_REL_TYPE)) throw new Error('Validation failed: comments rel missing');
    }
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

async function extractDocumentParagraphs(zip) {
    const documentXml = await zip.file('word/document.xml')?.async('string');
    if (!documentXml) throw new Error('word/document.xml not found');
    const xmlDoc = parseXmlStrict(documentXml, 'word/document.xml');
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

// ══════════════════════════════════════════════════════
// ── NEW: Gemini Chat Engine ──────────────────────────
// ══════════════════════════════════════════════════════

function buildSystemInstruction(paragraphs) {
    const listing = paragraphs.map(p => `[P${p.index}] ${p.text}`).join('\n');
    return [
        'You are a contract review AI assistant. The user has uploaded a document.',
        'Below is the document content. Each line is ONE SEPARATE PARAGRAPH, prefixed with [P#]:',
        '',
        listing,
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

async function sendGeminiChat(userMessage, paragraphs, apiKey) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(apiKey)}`;

    // Build the request with system instruction and multi-turn history
    const systemInstruction = buildSystemInstruction(paragraphs);

    // Build contents array: history + new user message
    const contents = [
        ...chatHistory,
        { role: 'user', parts: [{ text: userMessage }] }
    ];

    const response = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            systemInstruction: { parts: [{ text: systemInstruction }] },
            contents,
            generationConfig: {
                temperature: 0.3,
                maxOutputTokens: 8000
            }
        })
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

async function applyChatOperations(zip, operations, author) {
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
            const step = await runOperation(documentXml, op, author, runtimeContext);
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

    try {
        const result = await sendGeminiChat(userText, documentParagraphs, apiKey);

        // Remove thinking indicator
        thinkingEl.remove();

        let assistantHtml = escapeHtml(result.explanation).replace(/\n/g, '<br>');

        if (result.operations.length > 0) {
            addMsg('system', `Applying ${result.operations.length} operation(s)…`);
            const author = authorInput.value.trim() || 'Browser Demo AI';
            const opResults = await applyChatOperations(currentZip, result.operations, author);
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
