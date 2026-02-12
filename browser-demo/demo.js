import JSZip from 'https://esm.sh/jszip@3.10.1';
import {
    applyRedlineToOxml,
    applyHighlightToOoxml,
    injectCommentsIntoOoxml,
    configureLogger
} from '../src/taskpane/modules/reconciliation/standalone.js';

const DEMO_VERSION = '2026-02-12-chat-target-ref';
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

// ── State ──────────────────────────────────────────────
let currentZip = null;           // JSZip instance of the working document
let documentParagraphs = [];     // [{ index, text }] extracted from current docx
let chatHistory = [];            // Gemini multi-turn history [{ role, parts }]
let operationCount = 0;          // total operations applied across turns

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
    const textNodes = paragraph.getElementsByTagNameNS('*', 't');
    let text = '';
    for (const t of Array.from(textNodes)) text += t.textContent || '';
    return text;
}

function normalizeWhitespace(s) {
    return s.replace(/\s+/g, ' ').trim();
}

function getDocumentParagraphNodes(xmlDoc) {
    const body = getBodyElement(xmlDoc);
    if (body) return Array.from(body.getElementsByTagNameNS(NS_W, 'p'));
    return Array.from(xmlDoc.getElementsByTagNameNS('*', 'p'));
}

function findParagraphByStrictText(xmlDoc, targetText) {
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    const normalizedTarget = String(targetText || '').trim();
    if (!normalizedTarget) return null;

    const exact = paragraphs.find(p => getParagraphText(p).trim() === normalizedTarget);
    if (exact) return exact;

    const normTarget = normalizeWhitespace(normalizedTarget);
    return paragraphs.find(p => normalizeWhitespace(getParagraphText(p)) === normTarget) || null;
}

function findParagraphByExactText(xmlDoc, targetText) {
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    const normalizedTarget = String(targetText || '').trim();
    if (!normalizedTarget) return null;
    const strictMatch = findParagraphByStrictText(xmlDoc, normalizedTarget);
    if (strictMatch) return strictMatch;
    const normTarget = normalizeWhitespace(normalizedTarget);

    // 1. Target starts with paragraph text (Gemini may have merged multiple paragraphs)
    //    Find the first paragraph whose full text is a prefix of the target
    const startsWithMatch = paragraphs.find(p => {
        const pText = normalizeWhitespace(getParagraphText(p));
        return pText.length > 10 && normTarget.startsWith(pText);
    });
    if (startsWithMatch) {
        log(`[Fuzzy] Prefix match (target starts with paragraph): "${getParagraphText(startsWithMatch).trim().slice(0, 60)}…"`);
        return startsWithMatch;
    }

    // 2. Paragraph text contains the target or target contains paragraph text
    const containsMatch = paragraphs.find(p => {
        const pText = normalizeWhitespace(getParagraphText(p));
        return pText.length > 15 && normTarget.includes(pText);
    });
    if (containsMatch) {
        log(`[Fuzzy] Contains match: "${getParagraphText(containsMatch).trim().slice(0, 60)}…"`);
        return containsMatch;
    }

    // 3. Best overlap — score each paragraph by shared word count
    let bestScore = 0;
    let bestParagraph = null;
    const targetWords = new Set(normTarget.toLowerCase().split(/\s+/).filter(w => w.length > 2));
    for (const p of paragraphs) {
        const pText = getParagraphText(p).trim();
        if (!pText) continue;
        const pWords = normalizeWhitespace(pText).toLowerCase().split(/\s+/).filter(w => w.length > 2);
        const overlap = pWords.filter(w => targetWords.has(w)).length;
        const score = overlap / Math.max(targetWords.size, 1);
        if (score > bestScore && score > 0.5) {
            bestScore = score;
            bestParagraph = p;
        }
    }
    if (bestParagraph) {
        log(`[Fuzzy] Best word-overlap match (${(bestScore * 100).toFixed(0)}%): "${getParagraphText(bestParagraph).trim().slice(0, 60)}…"`);
        return bestParagraph;
    }

    return null;
}

function parseParagraphReference(rawValue) {
    if (rawValue == null) return null;
    if (typeof rawValue === 'number' && Number.isInteger(rawValue) && rawValue > 0) return rawValue;
    const text = String(rawValue).trim();
    if (!text) return null;
    const prefixed = text.match(/^\[?P(\d+)(?:\.\d+)?\]?$/i);
    if (prefixed) return Number.parseInt(prefixed[1], 10);
    const numeric = text.match(/^(\d+)$/);
    if (numeric) return Number.parseInt(numeric[1], 10);
    return null;
}

function stripLeadingParagraphMarker(text) {
    if (text == null) return '';
    return String(text).replace(/^\s*\[P\d+(?:\.\d+)?\]\s*/i, '').trim();
}

function splitLeadingParagraphMarker(text) {
    const raw = String(text || '');
    const marker = raw.match(/^\s*\[P(\d+)(?:\.\d+)?\]\s*/i);
    if (!marker) return { text: raw.trim(), targetRef: null };
    return {
        text: raw.replace(/^\s*\[P\d+(?:\.\d+)?\]\s*/i, '').trim(),
        targetRef: Number.parseInt(marker[1], 10)
    };
}

function findParagraphByReference(xmlDoc, targetRef) {
    if (!Number.isInteger(targetRef) || targetRef < 1) return null;
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    return paragraphs[targetRef - 1] || null;
}

function resolveTargetParagraph(xmlDoc, targetText, targetRef, opType) {
    const cleanTargetText = String(targetText || '').trim();
    const parsedRef = parseParagraphReference(targetRef);

    if (cleanTargetText) {
        const strictMatch = findParagraphByStrictText(xmlDoc, cleanTargetText);
        if (strictMatch) return { paragraph: strictMatch, resolvedBy: 'strict_text' };
    }

    if (parsedRef) {
        const byRef = findParagraphByReference(xmlDoc, parsedRef);
        if (byRef) {
            if (cleanTargetText) {
                const byRefText = getParagraphText(byRef).trim();
                if (normalizeWhitespace(byRefText) !== normalizeWhitespace(cleanTargetText)) {
                    log(`[Target] Using [P${parsedRef}] fallback for ${opType}; target text drifted.`);
                }
            } else {
                log(`[Target] Using [P${parsedRef}] fallback for ${opType}.`);
            }
            return { paragraph: byRef, resolvedBy: 'ref' };
        }
    }

    if (cleanTargetText) {
        const fuzzyMatch = findParagraphByExactText(xmlDoc, cleanTargetText);
        if (fuzzyMatch) return { paragraph: fuzzyMatch, resolvedBy: 'fuzzy_text' };
    }

    if (cleanTargetText) throw new Error(`Target paragraph not found: "${cleanTargetText}"`);
    if (parsedRef) throw new Error(`Target paragraph reference not found: [P${parsedRef}]`);
    throw new Error('Operation target missing: provide "target" text or "targetRef" ([P#]).');
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

// ── Apply operations (per-paragraph) ───────────────────
async function applyToParagraphByExactText(documentXml, targetText, modifiedText, author, targetRef = null) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'redline');
    const targetParagraph = resolved.paragraph;
    const currentParagraphText = getParagraphText(targetParagraph).trim();
    const paragraphXml = serializer.serializeToString(targetParagraph);
    const result = await applyRedlineToOxml(paragraphXml, currentParagraphText || targetText, modifiedText, { author, generateRedlines: true });
    if (!result?.hasChanges) return { documentXml, hasChanges: false, numberingXml: null };
    if (result.useNativeApi && !result.oxml) {
        const warning = 'Format-only fallback requires native Word API; browser demo skipped this operation.';
        log(`[WARN] ${warning}`);
        return { documentXml, hasChanges: false, numberingXml: null, warnings: [warning] };
    }
    if (typeof result.oxml !== 'string') {
        throw new Error('Reconciliation engine did not return OOXML for a changed redline operation');
    }
    const { replacementNodes, numberingXml } = extractReplacementNodes(result.oxml);
    const parent = targetParagraph.parentNode;
    for (const node of replacementNodes) parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);
    return { documentXml: serializer.serializeToString(xmlDoc), hasChanges: true, numberingXml };
}

async function applyHighlightToParagraphByExactText(documentXml, targetText, textToHighlight, color, author, targetRef = null) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'highlight');
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

async function applyCommentToParagraphByExactText(documentXml, targetText, textToComment, commentContent, author, targetRef = null) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');
    const resolved = resolveTargetParagraph(xmlDoc, targetText, targetRef, 'comment');
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

async function runOperation(documentXml, op, author) {
    if (op.type === 'highlight') return applyHighlightToParagraphByExactText(documentXml, op.target, op.textToHighlight, op.color, author, op.targetRef);
    if (op.type === 'comment') return applyCommentToParagraphByExactText(documentXml, op.target, op.textToComment, op.commentContent, author, op.targetRef);
    return applyToParagraphByExactText(documentXml, op.target, op.modified, author, op.targetRef);
}

// ── Package artifact helpers ───────────────────────────
async function ensureNumberingArtifacts(zip, numberingXml) {
    if (!numberingXml) return;
    const existing = await zip.file('word/numbering.xml')?.async('string');
    if (!existing) { log('[Demo] Adding numbering.xml'); zip.file('word/numbering.xml', numberingXml); }
    else { log('[Demo] Existing numbering.xml; keeping original'); }

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
        '',
        'CRITICAL TARGETING RULES:',
        '- Each [P#] line above is a SEPARATE paragraph in the document.',
        '- Always include "targetRef" using the paragraph label (example: "P12").',
        '- "targetRef" must point to the same paragraph as "target".',
        '- "target" MUST be the EXACT text of ONE SINGLE [P#] paragraph. Copy it character-for-character.',
        '- NEVER include the [P#] prefix in ANY operation field. The [P#] prefix is only a reference label, NOT part of the actual text.',
        '- NEVER combine or concatenate text from multiple [P#] paragraphs into one target.',
        '- If you need to modify multiple paragraphs, create a SEPARATE operation for EACH paragraph.',
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
        '  - Tables: use markdown table syntax (e.g., "| Col1 | Col2 |\\n|---|---|\\n| val | val |")',
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
            const targetRef = explicitRef || splitTarget.targetRef || null;
            const target = splitTarget.text;

            const normalizedOp = { ...op, type, target, targetRef };
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
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(apiKey)}`;

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

    let capturedNumberingXml = null;
    const capturedCommentsXml = [];
    const results = [];

    for (const op of operations) {
        const targetRefLabel = op.targetRef ? `[P${op.targetRef}] ` : '';
        const label = `${op.type}: ${targetRefLabel}"${(op.target || '').slice(0, 50)}…"`;
        log(`Applying: ${label}`);
        try {
            const step = await runOperation(documentXml, op, author);
            documentXml = step.documentXml;
            if (step.numberingXml) capturedNumberingXml = step.numberingXml;
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

        addMsg('system success', `Document loaded: <strong>${documentParagraphs.length} paragraphs</strong> found. You can now ask the AI to review it.`);
        log(`[Demo] v${DEMO_VERSION} — loaded ${file.name} (${documentParagraphs.length} paragraphs)`);

        chatInput.disabled = false;
        sendBtn.disabled = false;
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

            assistantHtml += buildOpSummaryHtml(opResults);

            // Show download button
            downloadBtn.style.display = '';
            downloadBtn.disabled = false;
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

// Log panel toggle
logToggle.addEventListener('click', () => {
    const isVisible = logEl.style.display === 'block';
    logEl.style.display = isVisible ? 'none' : 'block';
    logToggle.textContent = isVisible ? '▶ Engine log' : '▼ Engine log';
});

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
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(apiKey)}`;
    const prompt = ['Rewrite the following text as a cleaner sentence for a professional document.', 'Return plain text only, no quotes, markdown, bullets, or explanation.', `Text: ${originalText}`].join('\n');
    const response = await fetch(endpoint, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }], generationConfig: { temperature: 0.4, maxOutputTokens: 80 } }) });
    if (!response.ok) throw new Error(`Gemini API ${response.status}: ${(await response.text()).slice(0, 300)}`);
    const payload = await response.json();
    const suggestion = (payload?.candidates?.[0]?.content?.parts?.map(p => p?.text || '').join(' ').trim() || '');
    if (!suggestion) throw new Error('Gemini returned no text suggestion');
    return suggestion;
}

async function generateGeminiToolAction(apiKey) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(apiKey)}`;
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

    let capturedNumberingXml = null;
    const capturedCommentsXml = [];
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
        try { step = await runOperation(documentXml, op, author); }
        catch (error) {
            const msg = error?.message || String(error);
            const isSurprise = op.label === 'Gemini Surprise Tool Action' || op.label === 'AI Surprise Fallback';
            if (!isSurprise || !msg.includes('Target paragraph not found')) throw error;
            log(`[WARN] ${msg}`);
            log('[WARN] Retrying on safe target.');
            step = await runOperation(documentXml, buildSurpriseFallbackOperation(op), author);
        }
        documentXml = step.documentXml;
        if (step.numberingXml) capturedNumberingXml = step.numberingXml;
        if (step.commentsXml) capturedCommentsXml.push(step.commentsXml);
        if (step.warnings?.length > 0) for (const w of step.warnings) log(`  warning: ${w}`);
        log(`  changed: ${step.hasChanges}`);
    }

    { const p = new DOMParser(), s = new XMLSerializer(), d = p.parseFromString(documentXml, 'application/xml'); normalizeBodySectionOrder(d); documentXml = s.serializeToString(d); }
    zip.file('word/document.xml', documentXml);
    await ensureNumberingArtifacts(zip, capturedNumberingXml);
    for (const cx of capturedCommentsXml) await ensureCommentsArtifacts(zip, cx);
    await validateOutputDocx(zip);
    return await zip.generateAsync({ type: 'blob' });
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
        const output = await runKitchenSink(file, author, geminiApiKey);
        const outputName = file.name.replace(/\.docx$/i, '') + '-kitchen-sink-demo.docx';
        downloadBlob(output, outputName);
        addMsg('system success', 'Kitchen-sink demo completed. Document downloaded.');
        log('Kitchen-sink demo completed successfully.');
    } catch (err) {
        addMsg('system error', `Kitchen-sink failed: ${escapeHtml(err.message || String(err))}`);
        log(`[FATAL] ${err.message || String(err)}`);
        console.error(err);
    } finally { runBtn.disabled = false; }
});
