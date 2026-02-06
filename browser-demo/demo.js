import JSZip from 'https://esm.sh/jszip@3.10.1';
import {
    applyRedlineToOxml,
    configureLogger
} from '../src/taskpane/modules/reconciliation/standalone.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NUMBERING_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
const NUMBERING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml';

const fileInput = document.getElementById('docxFile');
const authorInput = document.getElementById('authorInput');
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

    const targets = [
        'DEMO_TEXT_TARGET',
        'DEMO FORMAT TARGET',
        'DEMO_LIST_TARGET',
        'DEMO_TABLE_TARGET'
    ];

    for (const marker of targets) {
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

async function ensureNumberingArtifacts(zip, numberingXml) {
    if (!numberingXml) return;
    const existingNumberingText = await zip.file('word/numbering.xml')?.async('string');
    if (!existingNumberingText) {
        zip.file('word/numbering.xml', numberingXml);
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

async function runKitchenSink(inputFile, author) {
    const zip = await JSZip.loadAsync(await inputFile.arrayBuffer());
    const documentFile = zip.file('word/document.xml');
    if (!documentFile) {
        throw new Error('word/document.xml not found in .docx');
    }

    let documentXml = await documentFile.async('string');
    documentXml = ensureDemoTargets(documentXml);

    let capturedNumberingXml = null;

    const operations = [
        {
            label: 'Text Edit',
            target: 'DEMO_TEXT_TARGET',
            modified: 'DEMO_TEXT_TARGET rewritten with extra words from the browser demo.'
        },
        {
            label: 'Format-Only (bold + underline)',
            target: 'DEMO FORMAT TARGET',
            modified: '**DEMO** ++FORMAT++ TARGET'
        },
        {
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
            label: 'Table Creation',
            target: 'DEMO_TABLE_TARGET',
            modified: [
                '| Item | Owner | Status |',
                '|---|---|---|',
                '| Engine refactor | Platform | Done |',
                '| Browser demo | UX | In Progress |',
                '| Documentation | QA | Planned |'
            ].join('\n')
        }
    ];

    for (const op of operations) {
        log(`Running: ${op.label}`);
        const step = await applyToParagraphByExactText(documentXml, op.target, op.modified, author);
        documentXml = step.documentXml;
        if (step.numberingXml) {
            capturedNumberingXml = step.numberingXml;
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
        const output = await runKitchenSink(file, author);
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
