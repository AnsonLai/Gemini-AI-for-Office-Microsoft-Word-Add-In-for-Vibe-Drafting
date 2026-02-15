import '../setup-xml-provider.mjs';

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

import {
    applyRedlineToOxmlWithListFallback,
    getDocumentParagraphNodes,
    getParagraphText,
    planListInsertionOnlyEdit,
    createDynamicNumberingIdState,
    reserveNextNumberingId,
    mergeNumberingXmlBySchemaOrder
} from '../../src/taskpane/modules/reconciliation/standalone.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parseXmlStrict(xmlText, label) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(xmlText, 'application/xml');
    const parseError = xmlDoc.getElementsByTagName('parsererror')[0];
    if (parseError) {
        throw new Error(`[XML parse error] ${label}: ${parseError.textContent || 'Unknown'}`);
    }
    return xmlDoc;
}

function normalizeWhitespace(text) {
    return String(text || '').replace(/\s+/g, ' ').trim();
}

function findParagraphBySnippet(xmlDoc, snippet) {
    const target = normalizeWhitespace(snippet).toLowerCase();
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    const matches = paragraphs.filter(paragraph => {
        const paragraphText = normalizeWhitespace(getParagraphText(paragraph)).toLowerCase();
        return paragraphText.includes(target);
    });
    if (matches.length === 0) {
        throw new Error(`Paragraph not found for snippet: "${snippet}"`);
    }
    if (matches.length > 1) {
        throw new Error(`Snippet matched multiple paragraphs: "${snippet}"`);
    }
    return matches[0];
}

function getPartName(partElement) {
    return partElement.getAttribute('pkg:name') || partElement.getAttribute('name') || '';
}

function isSectionPropertiesElement(node) {
    return !!node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === 'sectPr';
}

function extractFromPackage(packageXml) {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const pkgDoc = parseXmlStrict(packageXml, 'pkg output');
    const parts = Array.from(pkgDoc.getElementsByTagNameNS('*', 'part'));
    const documentPart = parts.find(part => getPartName(part) === '/word/document.xml');
    if (!documentPart) throw new Error('Package output missing /word/document.xml part');

    const xmlData = documentPart.getElementsByTagNameNS('*', 'xmlData')[0];
    if (!xmlData) throw new Error('Package document part missing pkg:xmlData');
    const documentNode = Array.from(xmlData.childNodes).find(node => node.nodeType === 1);
    if (!documentNode) throw new Error('Package document part missing XML payload');

    const body = documentNode.getElementsByTagNameNS('*', 'body')[0];
    const replacementNodes = body
        ? Array.from(body.childNodes).filter(node => node.nodeType === 1 && !isSectionPropertiesElement(node))
        : [documentNode];

    const numberingPart = parts.find(part => getPartName(part) === '/word/numbering.xml');
    let numberingXml = null;
    if (numberingPart) {
        const numberingXmlData = numberingPart.getElementsByTagNameNS('*', 'xmlData')[0];
        const numberingNode = numberingXmlData
            ? Array.from(numberingXmlData.childNodes).find(node => node.nodeType === 1)
            : null;
        if (numberingNode) numberingXml = serializer.serializeToString(numberingNode);
    }

    return { replacementNodes, numberingXml };
}

function extractReplacementNodes(outputOxml) {
    if (typeof outputOxml !== 'string' || !outputOxml.trim()) {
        throw new Error('Reconciliation returned no OOXML');
    }

    const parser = new DOMParser();
    if (outputOxml.includes('<pkg:package')) return extractFromPackage(outputOxml);
    if (outputOxml.includes('<w:document')) {
        const doc = parser.parseFromString(outputOxml, 'application/xml');
        const body = doc.getElementsByTagNameNS('*', 'body')[0];
        const replacementNodes = body
            ? Array.from(body.childNodes).filter(node => node.nodeType === 1 && !isSectionPropertiesElement(node))
            : Array.from(doc.childNodes).filter(node => node.nodeType === 1);
        return { replacementNodes, numberingXml: null };
    }

    const wrapped = `<root xmlns:w="${NS_W}">${outputOxml}</root>`;
    const fragmentDoc = parser.parseFromString(wrapped, 'application/xml');
    const replacementNodes = Array.from(fragmentDoc.documentElement.childNodes).filter(node => node.nodeType === 1);
    return { replacementNodes, numberingXml: null };
}

function getAttributeFirst(element, names) {
    for (const name of names) {
        const value = element.getAttribute(name);
        if (value != null && value !== '') return value;
    }
    return null;
}

function getElementId(element, names) {
    const raw = getAttributeFirst(element, names);
    const parsed = Number.parseInt(String(raw || ''), 10);
    return Number.isFinite(parsed) ? parsed : null;
}

function setElementId(element, preferredName, idValue) {
    element.setAttribute(preferredName, String(idValue));
}

function setElementVal(element, value) {
    element.setAttribute('w:val', String(value));
}

function createNumberingIdState(numberingXml) {
    return createDynamicNumberingIdState(numberingXml || '', {
        minId: 1,
        maxPreferred: 32767
    });
}

function remapNumberingPayloadForDocument(numberingXml, replacementNodes, numberingIdState) {
    const serializer = new XMLSerializer();
    const numberingDoc = parseXmlStrict(numberingXml, 'incoming numbering payload');
    const abstractNumMap = new Map();
    const numIdMap = new Map();

    const abstractNums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'abstractNum'));
    for (const abstractNum of abstractNums) {
        const oldId = getElementId(abstractNum, ['w:abstractNumId', 'abstractNumId']);
        if (oldId == null) continue;
        const newId = reserveNextNumberingId(numberingIdState, 'abstract');
        if (newId == null) continue;
        abstractNumMap.set(oldId, newId);
        setElementId(abstractNum, 'w:abstractNumId', newId);
    }

    const nums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'num'));
    for (const num of nums) {
        const oldNumId = getElementId(num, ['w:numId', 'numId']);
        if (oldNumId == null) continue;
        const newNumId = reserveNextNumberingId(numberingIdState, 'num');
        if (newNumId == null) continue;
        numIdMap.set(oldNumId, newNumId);
        setElementId(num, 'w:numId', newNumId);

        const abstractNumIdNode = Array.from(num.getElementsByTagNameNS('*', 'abstractNumId'))[0] || null;
        if (abstractNumIdNode) {
            const oldAbsRef = getElementId(abstractNumIdNode, ['w:val', 'val']);
            if (oldAbsRef != null && abstractNumMap.has(oldAbsRef)) {
                setElementVal(abstractNumIdNode, abstractNumMap.get(oldAbsRef));
            }
        }
    }

    const clonedNodes = replacementNodes.map(node => node.cloneNode(true));
    for (const node of clonedNodes) {
        const numIdNodes = Array.from(node.getElementsByTagNameNS('*', 'numId'));
        for (const numIdNode of numIdNodes) {
            const oldNumRef = getElementId(numIdNode, ['w:val', 'val']);
            if (oldNumRef != null && numIdMap.has(oldNumRef)) {
                setElementVal(numIdNode, numIdMap.get(oldNumRef));
            }
        }
    }

    return {
        numberingXml: serializer.serializeToString(numberingDoc),
        replacementNodes: clonedNodes
    };
}

function mergeNumberingXml(existingNumberingXml, incomingNumberingXml) {
    return mergeNumberingXmlBySchemaOrder(existingNumberingXml, incomingNumberingXml);
}

function getDirectWordChild(element, localName) {
    if (!element) return null;
    return Array.from(element.childNodes || []).find(
        node => node && node.nodeType === 1 && node.namespaceURI === NS_W && node.localName === localName
    ) || null;
}

function readParagraphListBinding(paragraph) {
    const pPr = getDirectWordChild(paragraph, 'pPr');
    if (!pPr) return null;
    const numPr = getDirectWordChild(pPr, 'numPr');
    if (!numPr) return null;
    const numIdEl = getDirectWordChild(numPr, 'numId');
    if (!numIdEl) return null;
    const ilvlEl = getDirectWordChild(numPr, 'ilvl');
    const numId = getAttributeFirst(numIdEl, ['w:val', 'val']);
    const ilvlRaw = ilvlEl ? getAttributeFirst(ilvlEl, ['w:val', 'val']) : '0';
    const ilvl = Number.parseInt(String(ilvlRaw || '0'), 10);
    return {
        numId: String(numId || ''),
        ilvl: Number.isFinite(ilvl) ? ilvl : 0
    };
}

function assertParagraphListLevel(xmlDoc, snippet, expectedIlvl) {
    const paragraph = findParagraphBySnippet(xmlDoc, snippet);
    const binding = readParagraphListBinding(paragraph);
    assertCondition(!!binding, `Expected list binding for "${snippet}"`);
    assertCondition(binding.ilvl === expectedIlvl, `Expected ilvl=${expectedIlvl} for "${snippet}", got ${binding.ilvl}`);
    return binding;
}

function assertAllReferencedNumIdsExist(xmlDoc, numberingXml) {
    const referencedNumIds = new Set(
        Array.from(xmlDoc.getElementsByTagNameNS('*', 'numId'))
            .map(node => getAttributeFirst(node, ['w:val', 'val']))
            .filter(Boolean)
            .map(String)
    );
    const numberingDoc = parseXmlStrict(numberingXml, 'merged numbering');
    const definedNumIds = new Set(
        Array.from(numberingDoc.getElementsByTagNameNS('*', 'num'))
            .map(node => getAttributeFirst(node, ['w:numId', 'numId']))
            .filter(Boolean)
            .map(String)
    );
    const missing = Array.from(referencedNumIds).filter(id => !definedNumIds.has(id));
    assertCondition(missing.length === 0, `Document references undefined numId(s): ${missing.join(', ')}`);
}

function assertStartOverrideForNumId(numberingXml, numId, expectedStart) {
    const numberingDoc = parseXmlStrict(numberingXml, 'merged numbering for startOverride checks');
    const nums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'num'));
    const targetNum = nums.find(node => {
        const id = getAttributeFirst(node, ['w:numId', 'numId']);
        return String(id || '') === String(numId);
    });
    assertCondition(!!targetNum, `Missing <w:num> definition for numId ${numId}`);

    const lvlOverrides = Array.from(targetNum.getElementsByTagNameNS('*', 'lvlOverride'));
    const lvl0Override = lvlOverrides.find(node => {
        const ilvl = getAttributeFirst(node, ['w:ilvl', 'ilvl']);
        return String(ilvl || '0') === '0';
    });
    assertCondition(!!lvl0Override, `Missing lvlOverride ilvl=0 for numId ${numId}`);

    const startOverride = Array.from(lvl0Override.getElementsByTagNameNS('*', 'startOverride'))[0] || null;
    assertCondition(!!startOverride, `Missing startOverride for numId ${numId}`);
    const startValue = getAttributeFirst(startOverride, ['w:val', 'val']);
    assertCondition(String(startValue || '') === String(expectedStart), `Expected startOverride ${expectedStart} for numId ${numId}, got ${startValue}`);
}

function assertNumFmtForNumId(numberingXml, numId, expectedNumFmt) {
    const numberingDoc = parseXmlStrict(numberingXml, 'merged numbering for numFmt checks');
    const nums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'num'));
    const targetNum = nums.find(node => {
        const id = getAttributeFirst(node, ['w:numId', 'numId']);
        return String(id || '') === String(numId);
    });
    assertCondition(!!targetNum, `Missing <w:num> definition for numId ${numId}`);

    const abstractNumIdNode = Array.from(targetNum.getElementsByTagNameNS('*', 'abstractNumId'))[0] || null;
    assertCondition(!!abstractNumIdNode, `Missing abstractNumId for numId ${numId}`);
    const abstractNumId = getAttributeFirst(abstractNumIdNode, ['w:val', 'val']);
    assertCondition(!!abstractNumId, `Invalid abstractNumId for numId ${numId}`);

    const abstractNums = Array.from(numberingDoc.getElementsByTagNameNS('*', 'abstractNum'));
    const targetAbstract = abstractNums.find(node => {
        const id = getAttributeFirst(node, ['w:abstractNumId', 'abstractNumId']);
        return String(id || '') === String(abstractNumId);
    });
    assertCondition(!!targetAbstract, `Missing abstractNum ${abstractNumId} for numId ${numId}`);

    const lvl0 = Array.from(targetAbstract.getElementsByTagNameNS('*', 'lvl'))
        .find(node => String(getAttributeFirst(node, ['w:ilvl', 'ilvl']) || '0') === '0');
    assertCondition(!!lvl0, `Missing lvl0 definition for numId ${numId}`);
    const numFmtNode = Array.from(lvl0.getElementsByTagNameNS('*', 'numFmt'))[0] || null;
    assertCondition(!!numFmtNode, `Missing numFmt for numId ${numId}`);
    const actualNumFmt = getAttributeFirst(numFmtNode, ['w:val', 'val']);
    assertCondition(String(actualNumFmt || '') === String(expectedNumFmt), `Expected numFmt ${expectedNumFmt} for numId ${numId}, got ${actualNumFmt}`);
}

function assertNumberingSchemaOrder(numberingXml) {
    const numberingDoc = parseXmlStrict(numberingXml, 'merged numbering for schema-order checks');
    const root = numberingDoc.documentElement;
    let sawNum = false;
    for (const child of Array.from(root.childNodes || [])) {
        if (!child || child.nodeType !== 1 || child.namespaceURI !== NS_W) continue;
        if (child.localName === 'num') {
            sawNum = true;
            continue;
        }
        if (child.localName === 'abstractNum' && sawNum) {
            throw new Error('Invalid numbering order: abstractNum appears after num');
        }
    }
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

function buildInsertedListParagraph(xmlDoc, anchorParagraph, entry, revisionId, author, dateIso) {
    const paragraph = xmlDoc.createElementNS(NS_W, 'w:p');

    const anchorPPr = getDirectWordChild(anchorParagraph, 'pPr');
    if (anchorPPr) {
        paragraph.appendChild(anchorPPr.cloneNode(true));
    }
    ensureListProperties(xmlDoc, paragraph, entry.ilvl, entry.numId);

    const ins = xmlDoc.createElementNS(NS_W, 'w:ins');
    ins.setAttribute('w:id', String(revisionId));
    ins.setAttribute('w:author', author);
    ins.setAttribute('w:date', dateIso);

    const run = xmlDoc.createElementNS(NS_W, 'w:r');
    const textNode = xmlDoc.createElementNS(NS_W, 'w:t');
    const safeText = String(entry.text || '').trim();
    if (/^\s|\s$/.test(safeText)) textNode.setAttribute('xml:space', 'preserve');
    textNode.textContent = safeText;
    run.appendChild(textNode);
    ins.appendChild(run);
    paragraph.appendChild(ins);

    return paragraph;
}

function assertCondition(condition, message) {
    if (!condition) throw new Error(message);
}

async function applyHeaderConversion(xmlDoc, serializer, numberingState, capturedNumberingXml, snippet) {
    const targetParagraph = findParagraphBySnippet(xmlDoc, snippet);
    const originalText = getParagraphText(targetParagraph).trim();
    const paragraphXml = serializer.serializeToString(targetParagraph);

    const result = await applyRedlineToOxmlWithListFallback(
        paragraphXml,
        originalText,
        originalText,
        {
            author: 'Regression Test',
            generateRedlines: true,
            listFallbackAllowExistingList: false,
            preferListStructuralFallback: true
        }
    );

    assertCondition(result?.hasChanges === true, `Header conversion did not apply for "${snippet}"`);

    const extracted = extractReplacementNodes(result.oxml);
    let replacementNodes = extracted.replacementNodes;
    let numberingXml = extracted.numberingXml || result.listStructuralFallbackNumberingXml || null;
    if (numberingXml) {
        const normalized = remapNumberingPayloadForDocument(numberingXml, replacementNodes, numberingState);
        replacementNodes = normalized.replacementNodes;
        numberingXml = normalized.numberingXml;
        capturedNumberingXml.push(numberingXml);
    }

    const firstReplacementParagraph = replacementNodes.find(
        node => node && node.nodeType === 1 && node.localName === 'p'
    ) || null;
    const replacementBinding = firstReplacementParagraph
        ? readParagraphListBinding(firstReplacementParagraph)
        : null;
    assertCondition(!!replacementBinding, `Converted header missing list binding for "${snippet}"`);
    assertCondition(replacementBinding.ilvl === 0, `Converted header ilvl must be 0 for "${snippet}"`);

    const parent = targetParagraph.parentNode;
    for (const node of replacementNodes) {
        parent.insertBefore(xmlDoc.importNode(node, true), targetParagraph);
    }
    parent.removeChild(targetParagraph);
    return replacementBinding;
}

function applyArchivalNestedInsertion(xmlDoc) {
    const targetSnippet = 'This copy is to be used solely for archival purposes to ensure compliance and record-keeping.';
    const targetParagraph = findParagraphBySnippet(xmlDoc, targetSnippet);
    const currentText = getParagraphText(targetParagraph).trim();
    const modifiedText = `${currentText}\n  2.2.1. Specifically, such retention must be legally required by the SEC or FCC.`;

    const insertionPlan = planListInsertionOnlyEdit(targetParagraph, modifiedText, {
        currentParagraphText: currentText
    });

    assertCondition(insertionPlan && insertionPlan.entries.length > 0, 'Archival insertion-only list plan was not produced');

    const parent = targetParagraph.parentNode;
    const insertionPoint = targetParagraph.nextSibling;
    const dateIso = new Date().toISOString();
    let revisionId = getNextTrackedChangeId(xmlDoc);
    for (const entry of insertionPlan.entries) {
        const listParagraph = buildInsertedListParagraph(
            xmlDoc,
            targetParagraph,
            { ...entry, numId: insertionPlan.numId },
            revisionId++,
            'Regression Test',
            dateIso
        );
        parent.insertBefore(listParagraph, insertionPoint);
    }
}

async function main() {
    const projectRoot = path.resolve(__dirname, '..', '..');
    const configuredSourceFolder = process.env.LIST_REGRESSION_SOURCE_FOLDER
        ? path.resolve(projectRoot, process.env.LIST_REGRESSION_SOURCE_FOLDER)
        : null;
    const sampleFolder = configuredSourceFolder || path.resolve(projectRoot, 'tests', 'sample_doc');
    const tmpRoot = path.resolve(projectRoot, 'tests', 'word-desktop', '.tmp');
    const workFolder = path.resolve(tmpRoot, 'list-regression-work');
    const pathsManifestPath = path.resolve(tmpRoot, 'list-regression-paths.json');
    const sourceDocumentXmlPath = path.resolve(sampleFolder, 'word', 'document.xml');
    const sourceNumberingXmlPath = path.resolve(sampleFolder, 'word', 'numbering.xml');

    if (!fs.existsSync(sourceDocumentXmlPath) || !fs.existsSync(sourceNumberingXmlPath)) {
        throw new Error(`Invalid regression source folder: ${sampleFolder}`);
    }

    fs.rmSync(workFolder, { recursive: true, force: true });
    fs.mkdirSync(tmpRoot, { recursive: true });
    fs.cpSync(sampleFolder, workFolder, { recursive: true });

    const documentXmlPath = path.resolve(workFolder, 'word', 'document.xml');
    const numberingXmlPath = path.resolve(workFolder, 'word', 'numbering.xml');

    const documentXml = fs.readFileSync(documentXmlPath, 'utf8');
    const numberingXml = fs.readFileSync(numberingXmlPath, 'utf8');

    const serializer = new XMLSerializer();
    const xmlDoc = parseXmlStrict(documentXml, 'word/document.xml');
    const numberingState = createNumberingIdState(numberingXml);
    const capturedNumberingXml = [];
    const headerBindings = [];

    const headerSnippets = [
        '1. DEFINITION OF CONFIDENTIAL INFORMATION',
        '2. EXCLUSIONS',
        '3. OBLIGATIONS OF RECEIVING PARTY',
        '4. TERM',
        '5. REQUIRED DISCLOSURE',
        '6. RETURN OF INFORMATION',
        '7. REMEDIES',
        '8. GOVERNING LAW',
        '9. GENERAL PROVISIONS'
    ];

    for (const snippet of headerSnippets) {
        const binding = await applyHeaderConversion(xmlDoc, serializer, numberingState, capturedNumberingXml, snippet);
        headerBindings.push(binding);
    }
    applyArchivalNestedInsertion(xmlDoc);

    fs.writeFileSync(documentXmlPath, serializer.serializeToString(xmlDoc), 'utf8');

    let mergedNumbering = numberingXml;
    for (const incoming of capturedNumberingXml) {
        mergedNumbering = mergeNumberingXml(mergedNumbering, incoming);
    }
    fs.writeFileSync(numberingXmlPath, mergedNumbering, 'utf8');

    // OOXML-level assertions that mirror known Word failure modes.
    assertAllReferencedNumIdsExist(xmlDoc, mergedNumbering);
    const exclusion1 = assertParagraphListLevel(xmlDoc, 'is or becomes generally available to the public', 0);
    const obligation1 = assertParagraphListLevel(xmlDoc, 'Use the Confidential Information solely for the Purpose', 0);
    const archivalNested = assertParagraphListLevel(xmlDoc, 'Specifically, such retention must be legally required by the SEC or FCC.', 2);
    assertCondition(headerBindings.length === 9, 'Expected nine converted header bindings.');
    assertCondition(headerBindings[1].numId !== exclusion1.numId, 'Header #2 must not share numId with exclusions sub-list.');
    assertCondition(headerBindings[2].numId !== exclusion1.numId, 'Header #3 must not share numId with exclusions sub-list.');
    assertCondition(headerBindings[2].numId !== obligation1.numId, 'Header #3 must not share numId with obligations sub-list.');
    assertCondition(archivalNested.numId.length > 0, 'Nested archival insertion has invalid list binding.');
    for (const binding of headerBindings) {
        assertNumFmtForNumId(mergedNumbering, binding.numId, 'decimal');
    }
    assertNumberingSchemaOrder(mergedNumbering);
    assertStartOverrideForNumId(mergedNumbering, headerBindings[0].numId, 1);
    assertStartOverrideForNumId(mergedNumbering, headerBindings[1].numId, 2);
    assertStartOverrideForNumId(mergedNumbering, headerBindings[2].numId, 3);
    assertStartOverrideForNumId(mergedNumbering, headerBindings[3].numId, 4);
    assertStartOverrideForNumId(mergedNumbering, headerBindings[4].numId, 5);
    assertStartOverrideForNumId(mergedNumbering, headerBindings[5].numId, 6);
    assertStartOverrideForNumId(mergedNumbering, headerBindings[6].numId, 7);
    assertStartOverrideForNumId(mergedNumbering, headerBindings[7].numId, 8);
    assertStartOverrideForNumId(mergedNumbering, headerBindings[8].numId, 9);

    const pathsManifest = {
        sourceFolder: sampleFolder,
        workFolder,
        outputDocx: path.resolve(tmpRoot, 'list-regression-output.docx'),
        outputInspectorJson: path.resolve(tmpRoot, 'list-regression-inspector.json')
    };
    fs.writeFileSync(pathsManifestPath, JSON.stringify(pathsManifest, null, 2), 'utf8');

    console.log('PASS: Built regression work package.');
    console.log(`Source folder: ${sampleFolder}`);
    console.log(`Work folder: ${workFolder}`);
    console.log(`Manifest: ${pathsManifestPath}`);
}

main().catch(error => {
    console.error('FAIL:', error.message || String(error));
    process.exit(1);
});
