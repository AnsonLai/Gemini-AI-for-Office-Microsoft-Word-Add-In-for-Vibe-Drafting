import JSZip from 'https://esm.sh/jszip@3.10.1';
import {
    applyRedlineToOxml,
    applyHighlightToOoxml,
    injectCommentsIntoOoxml,
    configureLogger
} from '../src/taskpane/modules/reconciliation/standalone.js';

const DEMO_VERSION = '2026-02-06-8';
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

const fileInput = document.getElementById('docxFile');
const authorInput = document.getElementById('authorInput');
const geminiApiKeyInput = document.getElementById('geminiApiKeyInput');
const saveGeminiKeyBtn = document.getElementById('saveGeminiKeyBtn');
const runBtn = document.getElementById('runBtn');
const statusEl = document.getElementById('status');
const logEl = document.getElementById('log');

function log(message) {
    logEl.textContent += `${message}\n`;
    logEl.scrollTop = logEl.scrollHeight;
}

function setStatus(message) {
    statusEl.textContent = message;
}

function getStoredGeminiApiKey() {
    try {
        return localStorage.getItem(GEMINI_API_KEY_STORAGE_KEY) || '';
    } catch (_error) {
        return '';
    }
}

function setStoredGeminiApiKey(apiKey) {
    try {
        if (apiKey) {
            localStorage.setItem(GEMINI_API_KEY_STORAGE_KEY, apiKey);
        } else {
            localStorage.removeItem(GEMINI_API_KEY_STORAGE_KEY);
        }
        return true;
    } catch (_error) {
        return false;
    }
}

async function generateGeminiRedlineSuggestion(originalText, apiKey) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(apiKey)}`;
    const prompt = [
        'Rewrite the following text as a cleaner sentence for a professional document.',
        'Return plain text only, no quotes, markdown, bullets, or explanation.',
        `Text: ${originalText}`
    ].join('\n');

    const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            contents: [
                {
                    role: 'user',
                    parts: [{ text: prompt }]
                }
            ],
            generationConfig: {
                temperature: 0.4,
                maxOutputTokens: 80
            }
        })
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Gemini API ${response.status}: ${errorText.slice(0, 300)}`);
    }

    const payload = await response.json();
    const suggestion = (
        payload?.candidates?.[0]?.content?.parts?.map(part => part?.text || '').join(' ').trim() || ''
    );

    if (!suggestion) {
        throw new Error('Gemini returned no text suggestion');
    }

    return suggestion;
}

function extractJsonObject(text) {
    if (!text) {
        throw new Error('No Gemini tool action text to parse');
    }

    const trimmed = text.trim();
    try {
        return JSON.parse(trimmed);
    } catch (_ignore) {
        // Try fallback parsing paths below.
    }

    const fencedMatch = trimmed.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
    if (fencedMatch?.[1]) {
        return JSON.parse(fencedMatch[1]);
    }

    const firstBrace = trimmed.indexOf('{');
    const lastBrace = trimmed.lastIndexOf('}');
    if (firstBrace >= 0 && lastBrace > firstBrace) {
        return JSON.parse(trimmed.slice(firstBrace, lastBrace + 1));
    }

    throw new Error('Gemini tool action did not contain a JSON object');
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
        const textToComment = String(args.textToComment || defaultTokenForTarget(target)).trim() || defaultTokenForTarget(target);
        const commentContent = String(args.commentContent || 'Gemini surprise comment: please review this section.').trim()
            || 'Gemini surprise comment: please review this section.';
        return {
            type: 'comment',
            label: 'Gemini Surprise Tool Action',
            target,
            textToComment,
            commentContent: commentContent.slice(0, 220)
        };
    }

    if (tool === 'highlight') {
        const textToHighlight = String(args.textToHighlight || defaultTokenForTarget(target)).trim() || defaultTokenForTarget(target);
        const colorCandidate = String(args.color || '').trim().toLowerCase();
        const color = ALLOWED_HIGHLIGHT_COLORS.includes(colorCandidate) ? colorCandidate : 'yellow';
        return {
            type: 'highlight',
            label: 'Gemini Surprise Tool Action',
            target,
            textToHighlight,
            color
        };
    }

    if (tool === 'redline') {
        const modified = String(args.modified || `${target} refined by Gemini surprise tool.`).trim()
            || `${target} refined by Gemini surprise tool.`;
        return {
            type: 'redline',
            label: 'Gemini Surprise Tool Action',
            target,
            modified: modified.slice(0, 260)
        };
    }

    throw new Error(`Unsupported Gemini tool: "${tool}"`);
}

async function generateGeminiToolAction(apiKey) {
    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(apiKey)}`;
    const prompt = [
        'You are choosing one surprise action for a DOCX demo pipeline.',
        'Return JSON only, no markdown and no commentary.',
        `Allowed targets: ${DEMO_MARKERS.join(', ')}`,
        'Choose exactly one tool:',
        '- comment -> args: { "target": string, "textToComment": string, "commentContent": string }',
        '- highlight -> args: { "target": string, "textToHighlight": string, "color": string }',
        '- redline -> args: { "target": string, "modified": string }',
        'Keep args short and practical.',
        'Response schema:',
        '{ "tool": "comment|highlight|redline", "args": { ... } }'
    ].join('\n');

    const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            contents: [
                {
                    role: 'user',
                    parts: [{ text: prompt }]
                }
            ],
            generationConfig: {
                temperature: 0.9,
                maxOutputTokens: 200
            }
        })
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Gemini API ${response.status}: ${errorText.slice(0, 300)}`);
    }

    const payload = await response.json();
    const rawText = (
        payload?.candidates?.[0]?.content?.parts?.map(part => part?.text || '').join('\n').trim() || ''
    );
    if (!rawText) {
        throw new Error('Gemini returned no tool action');
    }

    const parsed = extractJsonObject(rawText);
    return normalizeGeminiToolAction(parsed);
}

function parseXmlStrict(xmlText, label) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        throw new Error(`[XML parse error] ${label}: ${parseError.textContent || 'Unknown parsererror'}`);
    }
    return xmlDoc;
}

function isSectionPropertiesElement(node) {
    return !!node
        && node.nodeType === 1
        && node.namespaceURI === NS_W
        && node.localName === 'sectPr';
}

function getBodyElement(xmlDoc) {
    return xmlDoc.getElementsByTagNameNS('*', 'body')[0] || null;
}

function getDirectSectionProperties(body) {
    for (const child of Array.from(body.childNodes)) {
        if (isSectionPropertiesElement(child)) {
            return child;
        }
    }
    return null;
}

function insertBodyElementBeforeSectPr(body, element) {
    const sectPr = getDirectSectionProperties(body);
    if (sectPr) {
        body.insertBefore(element, sectPr);
    } else {
        body.appendChild(element);
    }
}

function normalizeBodySectionOrder(xmlDoc) {
    const body = getBodyElement(xmlDoc);
    if (!body) return;

    const sectPr = getDirectSectionProperties(body);
    if (!sectPr) return;

    let cursor = sectPr.nextSibling;
    while (cursor) {
        const next = cursor.nextSibling;
        if (cursor.nodeType === 1) {
            body.insertBefore(cursor, sectPr);
        }
        cursor = next;
    }
}

async function validateOutputDocx(zip) {
    const documentXml = await zip.file('word/document.xml')?.async('string');
    if (!documentXml) {
        throw new Error('Validation failed: missing word/document.xml');
    }

    const documentDoc = parseXmlStrict(documentXml, 'word/document.xml');
    normalizeBodySectionOrder(documentDoc);
    const body = getBodyElement(documentDoc);
    if (!body) {
        throw new Error('Validation failed: word/document.xml has no w:body');
    }

    const directBodyElements = Array.from(body.childNodes).filter(n => n.nodeType === 1);
    const sectPrIndexes = directBodyElements
        .map((node, idx) => ({ node, idx }))
        .filter(({ node }) => isSectionPropertiesElement(node))
        .map(({ idx }) => idx);
    if (sectPrIndexes.length > 1) {
        throw new Error('Validation failed: multiple body-level w:sectPr elements');
    }
    if (sectPrIndexes.length === 1 && sectPrIndexes[0] !== directBodyElements.length - 1) {
        throw new Error('Validation failed: body-level w:sectPr is not the last body child');
    }

    const nestedParagraphs = documentDoc.getElementsByTagNameNS(NS_W, 'tc');
    for (const tc of Array.from(nestedParagraphs)) {
        const directCellChildren = Array.from(tc.childNodes).filter(n => n.nodeType === 1);
        for (const child of directCellChildren) {
            if (child.namespaceURI === NS_W && child.localName === 'p') {
                const nested = Array.from(child.childNodes).find(n =>
                    n.nodeType === 1 &&
                    n.namespaceURI === NS_W &&
                    n.localName === 'p'
                );
                if (nested) {
                    throw new Error('Validation failed: nested w:p found inside table cell paragraph');
                }
            }
        }
    }

    const hasNumberingUsage = documentDoc.getElementsByTagNameNS(NS_W, 'numPr').length > 0;
    const hasCommentUsage = (
        documentDoc.getElementsByTagNameNS(NS_W, 'commentRangeStart').length > 0 ||
        documentDoc.getElementsByTagNameNS(NS_W, 'commentRangeEnd').length > 0 ||
        documentDoc.getElementsByTagNameNS(NS_W, 'commentReference').length > 0
    );
    const numberingXml = await zip.file('word/numbering.xml')?.async('string');
    const commentsXml = await zip.file('word/comments.xml')?.async('string');
    const hasNumberingPart = !!numberingXml;
    const hasCommentsPart = !!commentsXml;

    if (hasNumberingPart) {
        parseXmlStrict(numberingXml, 'word/numbering.xml');
    } else if (hasNumberingUsage) {
        throw new Error('Validation failed: document uses numbering but word/numbering.xml is missing');
    }

    if (hasCommentsPart) {
        parseXmlStrict(commentsXml, 'word/comments.xml');
    } else if (hasCommentUsage) {
        throw new Error('Validation failed: document uses comments but word/comments.xml is missing');
    }

    const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');
    if (!contentTypesXml) {
        throw new Error('Validation failed: missing [Content_Types].xml');
    }
    const contentTypesDoc = parseXmlStrict(contentTypesXml, '[Content_Types].xml');

    const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string');
    if (!relsXml) {
        throw new Error('Validation failed: missing word/_rels/document.xml.rels');
    }
    const relsDoc = parseXmlStrict(relsXml, 'word/_rels/document.xml.rels');

    if (hasNumberingPart) {
        const overrides = Array.from(contentTypesDoc.getElementsByTagNameNS('*', 'Override'));
        const hasNumberingOverride = overrides.some(o =>
            (o.getAttribute('PartName') || '').toLowerCase() === '/word/numbering.xml'
                && (o.getAttribute('ContentType') || '') === NUMBERING_CONTENT_TYPE
        );
        if (!hasNumberingOverride) {
            throw new Error('Validation failed: numbering part exists but [Content_Types].xml override is missing');
        }

        const rels = Array.from(relsDoc.getElementsByTagNameNS('*', 'Relationship'));
        const hasNumberingRel = rels.some(r => (r.getAttribute('Type') || '') === NUMBERING_REL_TYPE);
        if (!hasNumberingRel) {
            throw new Error('Validation failed: numbering part exists but document relationship is missing');
        }
    }

    if (hasCommentsPart) {
        const overrides = Array.from(contentTypesDoc.getElementsByTagNameNS('*', 'Override'));
        const hasCommentsOverride = overrides.some(o =>
            (o.getAttribute('PartName') || '').toLowerCase() === '/word/comments.xml'
                && (o.getAttribute('ContentType') || '') === COMMENTS_CONTENT_TYPE
        );
        if (!hasCommentsOverride) {
            throw new Error('Validation failed: comments part exists but [Content_Types].xml override is missing');
        }

        const rels = Array.from(relsDoc.getElementsByTagNameNS('*', 'Relationship'));
        const hasCommentsRel = rels.some(r => (r.getAttribute('Type') || '') === COMMENTS_REL_TYPE);
        if (!hasCommentsRel) {
            throw new Error('Validation failed: comments part exists but document relationship is missing');
        }
    }
}

configureLogger({
    log: (...args) => log(args.map(String).join(' ')),
    warn: (...args) => log(`[WARN] ${args.map(String).join(' ')}`),
    error: (...args) => log(`[ERROR] ${args.map(String).join(' ')}`)
});

function getPartName(partElement) {
    return (
        partElement.getAttribute('pkg:name') ||
        partElement.getAttribute('name') ||
        ''
    );
}

function getParagraphText(paragraph) {
    const textNodes = paragraph.getElementsByTagNameNS('*', 't');
    let text = '';
    for (const t of Array.from(textNodes)) {
        text += t.textContent || '';
    }
    return text;
}

function findParagraphByExactText(xmlDoc, targetText) {
    const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS('*', 'p'));
    const normalizedTarget = targetText.trim();
    return paragraphs.find(p => getParagraphText(p).trim() === normalizedTarget) || null;
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

function ensureDemoTargets(documentXml) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');

    const body = getBodyElement(xmlDoc);
    if (!body) {
        throw new Error('Could not find w:body in document.xml');
    }

    for (const marker of DEMO_MARKERS) {
        if (!findParagraphByExactText(xmlDoc, marker)) {
            insertBodyElementBeforeSectPr(body, createSimpleParagraph(xmlDoc, marker));
            log(`Inserted missing marker paragraph: ${marker}`);
        }
    }

    normalizeBodySectionOrder(xmlDoc);

    return serializer.serializeToString(xmlDoc);
}

function extractFromPackage(packageXml) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const pkgDoc = parser.parseFromString(packageXml, 'application/xml');

    const parts = Array.from(pkgDoc.getElementsByTagNameNS('*', 'part'));
    const documentPart = parts.find(p => getPartName(p) === '/word/document.xml');
    if (!documentPart) {
        throw new Error('Package output missing /word/document.xml part');
    }

    const xmlData = documentPart.getElementsByTagNameNS('*', 'xmlData')[0];
    if (!xmlData) {
        throw new Error('Package document part missing pkg:xmlData');
    }

    const documentNode = Array.from(xmlData.childNodes).find(n => n.nodeType === 1);
    if (!documentNode) {
        throw new Error('Package document part missing XML payload');
    }

    const body = documentNode.getElementsByTagNameNS('*', 'body')[0];
    const replacementNodes = body
        ? Array.from(body.childNodes).filter(n => n.nodeType === 1 && !isSectionPropertiesElement(n))
        : [documentNode];

    const numberingPart = parts.find(p => getPartName(p) === '/word/numbering.xml');
    let numberingXml = null;
    if (numberingPart) {
        const numberingXmlData = numberingPart.getElementsByTagNameNS('*', 'xmlData')[0];
        const numberingNode = numberingXmlData
            ? Array.from(numberingXmlData.childNodes).find(n => n.nodeType === 1)
            : null;
        if (numberingNode) {
            numberingXml = serializer.serializeToString(numberingNode);
        }
    }

    return { replacementNodes, numberingXml };
}

function extractReplacementNodes(outputOxml) {
    const parser = new DOMParser();

    if (outputOxml.includes('<pkg:package')) {
        return extractFromPackage(outputOxml);
    }

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

async function applyToParagraphByExactText(documentXml, targetText, modifiedText, author) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');

    const targetParagraph = findParagraphByExactText(xmlDoc, targetText);
    if (!targetParagraph) {
        throw new Error(`Target paragraph not found: "${targetText}"`);
    }

    const paragraphXml = serializer.serializeToString(targetParagraph);
    const result = await applyRedlineToOxml(paragraphXml, targetText, modifiedText, {
        author,
        generateRedlines: true
    });

    if (!result.hasChanges) {
        return { documentXml, hasChanges: false, numberingXml: null };
    }

    const { replacementNodes, numberingXml } = extractReplacementNodes(result.oxml);
    const parent = targetParagraph.parentNode;

    for (const node of replacementNodes) {
        parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    }
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);

    return {
        documentXml: serializer.serializeToString(xmlDoc),
        hasChanges: true,
        numberingXml
    };
}

async function applyHighlightToParagraphByExactText(documentXml, targetText, textToHighlight, color, author) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');

    const targetParagraph = findParagraphByExactText(xmlDoc, targetText);
    if (!targetParagraph) {
        throw new Error(`Target paragraph not found for highlight: "${targetText}"`);
    }

    const paragraphXml = serializer.serializeToString(targetParagraph);
    const highlightedXml = applyHighlightToOoxml(paragraphXml, textToHighlight, color, {
        generateRedlines: true,
        author
    });

    if (!highlightedXml || highlightedXml === paragraphXml) {
        return { documentXml, hasChanges: false };
    }

    const { replacementNodes } = extractReplacementNodes(highlightedXml);
    const parent = targetParagraph.parentNode;
    for (const node of replacementNodes) {
        parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    }
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);

    return {
        documentXml: serializer.serializeToString(xmlDoc),
        hasChanges: true
    };
}

async function applyCommentToParagraphByExactText(documentXml, targetText, textToComment, commentContent, author) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const xmlDoc = parser.parseFromString(documentXml, 'application/xml');

    const targetParagraph = findParagraphByExactText(xmlDoc, targetText);
    if (!targetParagraph) {
        throw new Error(`Target paragraph not found for comment: "${targetText}"`);
    }

    const paragraphXml = serializer.serializeToString(targetParagraph);
    const commentResult = injectCommentsIntoOoxml(
        paragraphXml,
        [{
            paragraphIndex: 1,
            textToFind: textToComment,
            commentContent
        }],
        { author }
    );

    if (!commentResult.commentsApplied) {
        return {
            documentXml,
            hasChanges: false,
            commentsXml: null,
            warnings: commentResult.warnings || []
        };
    }

    const { replacementNodes } = extractReplacementNodes(commentResult.oxml);
    const parent = targetParagraph.parentNode;
    for (const node of replacementNodes) {
        parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    }
    parent.removeChild(targetParagraph);
    normalizeBodySectionOrder(xmlDoc);

    return {
        documentXml: serializer.serializeToString(xmlDoc),
        hasChanges: true,
        commentsXml: commentResult.commentsXml || null,
        warnings: commentResult.warnings || []
    };
}

async function runOperation(documentXml, op, author) {
    if (op.type === 'highlight') {
        return applyHighlightToParagraphByExactText(
            documentXml,
            op.target,
            op.textToHighlight,
            op.color,
            author
        );
    }

    if (op.type === 'comment') {
        return applyCommentToParagraphByExactText(
            documentXml,
            op.target,
            op.textToComment,
            op.commentContent,
            author
        );
    }

    return applyToParagraphByExactText(documentXml, op.target, op.modified, author);
}

function buildSurpriseFallbackOperation(op) {
    const safeTarget = 'DEMO FORMAT TARGET';

    if (op.type === 'highlight') {
        return {
            ...op,
            target: safeTarget,
            textToHighlight: 'FORMAT',
            color: ALLOWED_HIGHLIGHT_COLORS.includes(String(op.color || '').toLowerCase()) ? op.color : 'yellow'
        };
    }

    if (op.type === 'comment') {
        return {
            ...op,
            target: safeTarget,
            textToComment: 'FORMAT'
        };
    }

    return {
        ...op,
        target: safeTarget,
        modified: 'DEMO FORMAT TARGET updated by Gemini surprise retry.'
    };
}

async function ensureNumberingArtifacts(zip, numberingXml) {
    if (!numberingXml) return;
    const existingNumberingText = await zip.file('word/numbering.xml')?.async('string');
    if (!existingNumberingText) {
        log('[Demo] Adding numbering.xml from list-generation package output');
        zip.file('word/numbering.xml', numberingXml);
    } else {
        log('[Demo] Existing numbering.xml detected; keeping original numbering part');
    }

    const parser = new DOMParser();
    const serializer = new XMLSerializer();

    // [Content_Types].xml
    const contentTypesText = await zip.file('[Content_Types].xml')?.async('string');
    if (contentTypesText) {
        const ctDoc = parser.parseFromString(contentTypesText, 'application/xml');
        const overrides = Array.from(ctDoc.getElementsByTagNameNS('*', 'Override'));
        const hasNumberingOverride = overrides.some(o =>
            (o.getAttribute('PartName') || '').toLowerCase() === '/word/numbering.xml'
        );
        if (!hasNumberingOverride) {
            const override = ctDoc.createElementNS(NS_CT, 'Override');
            override.setAttribute('PartName', '/word/numbering.xml');
            override.setAttribute('ContentType', NUMBERING_CONTENT_TYPE);
            ctDoc.documentElement.appendChild(override);
            zip.file('[Content_Types].xml', serializer.serializeToString(ctDoc));
        }
    }

    // word/_rels/document.xml.rels
    const relsPath = 'word/_rels/document.xml.rels';
    const relsText = await zip.file(relsPath)?.async('string');
    if (relsText) {
        const relsDoc = parser.parseFromString(relsText, 'application/xml');
        const relsRoot = relsDoc.getElementsByTagNameNS('*', 'Relationships')[0] || relsDoc.documentElement;
        const rels = Array.from(relsRoot.getElementsByTagNameNS('*', 'Relationship'));
        const hasNumberingRel = rels.some(r => (r.getAttribute('Type') || '') === NUMBERING_REL_TYPE);

        if (!hasNumberingRel) {
            let max = 0;
            for (const rel of rels) {
                const id = rel.getAttribute('Id') || '';
                const n = parseInt(id.replace(/^rId/i, ''), 10);
                if (!Number.isNaN(n)) max = Math.max(max, n);
            }
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
    const existingCommentsText = await zip.file(commentsPath)?.async('string');

    if (!existingCommentsText) {
        log('[Demo] Adding comments.xml from comment operation');
        zip.file(commentsPath, commentsXml);
    } else {
        const existingDoc = parseXmlStrict(existingCommentsText, 'word/comments.xml (existing)');
        const incomingDoc = parseXmlStrict(commentsXml, 'word/comments.xml (incoming)');

        const existingRoot = existingDoc.documentElement;
        const incomingRoot = incomingDoc.documentElement;
        const existingIds = new Set(
            Array.from(existingRoot.getElementsByTagNameNS(NS_W, 'comment'))
                .map(c => c.getAttribute('w:id') || c.getAttribute('id'))
                .filter(Boolean)
        );

        for (const incomingComment of Array.from(incomingRoot.getElementsByTagNameNS(NS_W, 'comment'))) {
            const incomingId = incomingComment.getAttribute('w:id') || incomingComment.getAttribute('id');
            if (incomingId && existingIds.has(incomingId)) {
                throw new Error(`Validation failed: duplicate comment id detected while merging comments.xml (id=${incomingId})`);
            }
            existingRoot.appendChild(existingDoc.importNode(incomingComment, true));
        }

        zip.file(commentsPath, serializer.serializeToString(existingDoc));
    }

    // [Content_Types].xml
    const contentTypesText = await zip.file('[Content_Types].xml')?.async('string');
    if (contentTypesText) {
        const ctDoc = parser.parseFromString(contentTypesText, 'application/xml');
        const overrides = Array.from(ctDoc.getElementsByTagNameNS('*', 'Override'));
        const hasCommentsOverride = overrides.some(o =>
            (o.getAttribute('PartName') || '').toLowerCase() === '/word/comments.xml'
        );
        if (!hasCommentsOverride) {
            const override = ctDoc.createElementNS(NS_CT, 'Override');
            override.setAttribute('PartName', '/word/comments.xml');
            override.setAttribute('ContentType', COMMENTS_CONTENT_TYPE);
            ctDoc.documentElement.appendChild(override);
            zip.file('[Content_Types].xml', serializer.serializeToString(ctDoc));
        }
    }

    // word/_rels/document.xml.rels
    const relsPath = 'word/_rels/document.xml.rels';
    const relsText = await zip.file(relsPath)?.async('string');
    if (relsText) {
        const relsDoc = parser.parseFromString(relsText, 'application/xml');
        const relsRoot = relsDoc.getElementsByTagNameNS('*', 'Relationships')[0] || relsDoc.documentElement;
        const rels = Array.from(relsRoot.getElementsByTagNameNS('*', 'Relationship'));
        const hasCommentsRel = rels.some(r => (r.getAttribute('Type') || '') === COMMENTS_REL_TYPE);

        if (!hasCommentsRel) {
            let max = 0;
            for (const rel of rels) {
                const id = rel.getAttribute('Id') || '';
                const n = parseInt(id.replace(/^rId/i, ''), 10);
                if (!Number.isNaN(n)) max = Math.max(max, n);
            }
            const rel = relsDoc.createElementNS(NS_RELS, 'Relationship');
            rel.setAttribute('Id', `rId${max + 1}`);
            rel.setAttribute('Type', COMMENTS_REL_TYPE);
            rel.setAttribute('Target', 'comments.xml');
            relsRoot.appendChild(rel);
            zip.file(relsPath, serializer.serializeToString(relsDoc));
        }
    }
}

async function runKitchenSink(inputFile, author, geminiApiKey) {
    const zip = await JSZip.loadAsync(await inputFile.arrayBuffer());
    const documentFile = zip.file('word/document.xml');
    if (!documentFile) {
        throw new Error('word/document.xml not found in .docx');
    }

    let documentXml = await documentFile.async('string');
    parseXmlStrict(documentXml, 'word/document.xml (input)');
    documentXml = ensureDemoTargets(documentXml);

    let capturedNumberingXml = null;
    const capturedCommentsXml = [];
    const fallbackRedlineText = 'DEMO_TEXT_TARGET rewritten with extra words from the browser demo.';
    const fallbackToolOperation = {
        type: 'comment',
        label: 'AI Surprise Fallback',
        target: 'DEMO FORMAT TARGET',
        textToComment: 'FORMAT',
        commentContent: 'Fallback AI action: please review the formatting language here.'
    };
    let geminiRedlineText = fallbackRedlineText;
    let geminiToolOperation = fallbackToolOperation;

    if (geminiApiKey) {
        log('Generating Gemini redline suggestion for DEMO_TEXT_TARGET...');
        try {
            const suggested = await generateGeminiRedlineSuggestion('DEMO_TEXT_TARGET', geminiApiKey);
            if (suggested.trim() && suggested.trim() !== 'DEMO_TEXT_TARGET') {
                geminiRedlineText = suggested.trim();
                log(`Gemini suggestion selected: ${geminiRedlineText}`);
            } else {
                log('[WARN] Gemini suggestion matched source text; using fallback rewrite.');
            }
        } catch (error) {
            log(`[WARN] Gemini suggestion failed; using fallback rewrite. ${error.message || String(error)}`);
        }

        log('Generating Gemini surprise tool action...');
        try {
            const action = await generateGeminiToolAction(geminiApiKey);
            geminiToolOperation = action;
            log(`Gemini surprise action selected: ${action.type} on "${action.target}"`);
        } catch (error) {
            log(`[WARN] Gemini surprise action failed; using fallback action. ${error.message || String(error)}`);
        }
    } else {
        log('[WARN] No Gemini API key provided; using fallback rewrite and fallback action.');
    }

    const operations = [
        {
            type: 'redline',
            label: 'Text Edit',
            target: 'DEMO_TEXT_TARGET',
            modified: geminiRedlineText
        },
        {
            type: 'redline',
            label: 'Format-Only (bold + underline)',
            target: 'DEMO FORMAT TARGET',
            modified: '**DEMO** ++FORMAT++ TARGET'
        },
        {
            type: 'redline',
            label: 'Bullets + Sub-bullets',
            target: 'DEMO_LIST_TARGET',
            modified: [
                '- Browser demo top bullet',
                '  - Nested bullet A',
                '  - Nested bullet B',
                '- Browser demo second bullet'
            ].join('\n')
        },
        {
            type: 'redline',
            label: 'Table Creation',
            target: 'DEMO_TABLE_TARGET',
            modified: [
                '| Item | Owner | Status |',
                '|---|---|---|',
                '| Engine refactor | Platform | Done |',
                '| Browser demo | UX | In Progress |',
                '| Documentation | QA | Planned |'
            ].join('\n')
        },
        {
            ...geminiToolOperation
        }
    ];

    for (const op of operations) {
        log(`Running: ${op.label}`);
        let step;
        try {
            step = await runOperation(documentXml, op, author);
        } catch (error) {
            const message = error?.message || String(error);
            const isSurpriseAction = op.label === 'Gemini Surprise Tool Action' || op.label === 'AI Surprise Fallback';
            const isTargetMissing = message.includes('Target paragraph not found');

            if (!isSurpriseAction || !isTargetMissing) {
                throw error;
            }

            log(`[WARN] ${message}`);
            log('[WARN] Retrying surprise action on a safe target (DEMO FORMAT TARGET).');
            const retryOp = buildSurpriseFallbackOperation(op);
            step = await runOperation(documentXml, retryOp, author);
        }

        documentXml = step.documentXml;
        if (step.numberingXml) {
            capturedNumberingXml = step.numberingXml;
        }
        if (step.commentsXml) {
            capturedCommentsXml.push(step.commentsXml);
        }
        if (step.warnings && step.warnings.length > 0) {
            for (const warning of step.warnings) {
                log(`  warning: ${warning}`);
            }
        }
        log(`  changed: ${step.hasChanges}`);
    }

    {
        const parser = new DOMParser();
        const serializer = new XMLSerializer();
        const finalDoc = parser.parseFromString(documentXml, 'application/xml');
        normalizeBodySectionOrder(finalDoc);
        documentXml = serializer.serializeToString(finalDoc);
    }

    zip.file('word/document.xml', documentXml);
    await ensureNumberingArtifacts(zip, capturedNumberingXml);
    for (const commentsXml of capturedCommentsXml) {
        await ensureCommentsArtifacts(zip, commentsXml);
    }
    await validateOutputDocx(zip);

    const outputBlob = await zip.generateAsync({ type: 'blob' });
    return outputBlob;
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

runBtn.addEventListener('click', async () => {
    const file = fileInput.files?.[0];
    if (!file) {
        setStatus('Please choose a .docx file.');
        return;
    }

    runBtn.disabled = true;
    logEl.textContent = '';
    setStatus('Running kitchen-sink demo...');

    try {
        const author = authorInput.value.trim() || 'Browser Demo AI';
        const geminiApiKey = geminiApiKeyInput?.value.trim() || '';
        if (geminiApiKey) {
            setStoredGeminiApiKey(geminiApiKey);
        }
        log(`[Demo] Version: ${DEMO_VERSION}`);
        const output = await runKitchenSink(file, author, geminiApiKey);
        const outputName = file.name.replace(/\.docx$/i, '') + '-kitchen-sink-demo.docx';
        downloadBlob(output, outputName);
        setStatus('Done. Modified document downloaded.');
        log('Demo completed successfully.');
    } catch (err) {
        setStatus('Failed.');
        log(`[FATAL] ${err.message || String(err)}`);
        console.error(err);
    } finally {
        runBtn.disabled = false;
    }
});

saveGeminiKeyBtn.addEventListener('click', () => {
    const key = geminiApiKeyInput?.value.trim() || '';
    const saved = setStoredGeminiApiKey(key);

    if (!saved) {
        setStatus('Unable to save Gemini API key in this browser.');
        log('[WARN] Failed to persist Gemini API key to localStorage.');
        return;
    }

    if (key) {
        setStatus('Gemini API key saved in this browser.');
        log('Gemini API key saved.');
    } else {
        setStatus('Gemini API key cleared from this browser.');
        log('Gemini API key cleared.');
    }
});

if (geminiApiKeyInput) {
    const storedKey = getStoredGeminiApiKey();
    if (storedKey) {
        geminiApiKeyInput.value = storedKey;
    }
}
