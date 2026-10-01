import {
    getDocumentParagraphNodes,
    ingestWordOoxmlToMarkdown,
    openDocx,
    parseOoxmlSafe,
    serializeOoxml
} from '@ansonlai/docx-redline-js';

const WORD_MAIN_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const textDecoder = new TextDecoder('utf-8');

function decodePart(part) {
    if (typeof part === 'string') return part;
    if (part instanceof Uint8Array) return textDecoder.decode(part);
    return null;
}

function requiredDocumentXml(document) {
    const xml = decodePart(document.entries.get('word/document.xml'));
    if (!xml) throw new Error('word/document.xml not found');
    return xml;
}

function formatParagraphNode(paragraph) {
    const paragraphXml = serializeOoxml(paragraph);
    const wrappedDocument = `<w:document xmlns:w="${WORD_MAIN_NS}"><w:body>${paragraphXml}<w:sectPr/></w:body></w:document>`;
    return String(ingestWordOoxmlToMarkdown(wrappedDocument) || '').trim();
}

function stripProjectedListMarker(formattedText, list) {
    const text = String(formattedText || '').trim();
    if (!list) return text;
    if (list.format === 'bullet') return text.replace(/^(?:[-*+•▪‣]\s+)/, '');
    if (list.format === 'decimal') return text.replace(/^(?:\d+\.|[A-Za-z]\.)\s+/, '');
    return text;
}

function formatPromptParagraph(paragraph, formattedText) {
    if (!paragraph.list) return formattedText;
    const depth = Number.isInteger(paragraph.list.level) && paragraph.list.level > 0
        ? paragraph.list.level
        : 0;
    const marker = paragraph.list.format === 'bullet' ? '-' : String(paragraph.list.label || '').trim();
    const content = stripProjectedListMarker(formattedText, paragraph.list);
    if (!marker || !content) return formattedText;
    return `${'  '.repeat(depth)}${marker} ${content}`;
}

function markerSeedAnchorIsSafe(paragraph) {
    const text = String(paragraph?.exactText || '');
    return !paragraph?.inTable
        && !paragraph?.list
        && !!text.trim()
        && !/[\\*_`~\[\]#|]/.test(text)
        && !/[\r\n]/.test(text);
}

/**
 * Seeds standalone paragraph markers through the public structured-content
 * operation contract. It refuses to proceed unless source paragraphs and all
 * requested marker paragraphs are still visible after the facade write.
 */
export async function seedParagraphMarkers(session, markers, author) {
    if (!session || typeof session.inspect !== 'function' || typeof session.applyOperations !== 'function') {
        throw new TypeError('A document session is required to seed paragraph markers.');
    }
    if (!Array.isArray(markers) || markers.some(marker => typeof marker !== 'string' || !marker.trim())) {
        throw new TypeError('Paragraph markers must be non-empty strings.');
    }
    if (new Set(markers).size !== markers.length) {
        throw new TypeError('Paragraph markers must be unique.');
    }

    const before = session.inspect();
    if (before.status !== 'ok') throw new Error(before.error?.message || 'Could not inspect document before marker seeding.');
    const existingTexts = new Set(before.paragraphs.map(paragraph => paragraph.exactText));
    const missing = markers.filter(marker => !existingTexts.has(marker));
    if (missing.length === 0) return { added: [], inspection: before };

    let anchor = null;
    for (let index = before.paragraphs.length - 1; index >= 0; index -= 1) {
        if (markerSeedAnchorIsSafe(before.paragraphs[index])) {
            anchor = before.paragraphs[index];
            break;
        }
    }
    if (!anchor) throw new Error('Cannot seed marker paragraphs without a plain-text body paragraph outside tables and lists.');

    const operation = {
        type: 'redline',
        target: {
            index: anchor.index,
            exactText: anchor.exactText,
            ...(anchor.paragraphId ? { paragraphId: anchor.paragraphId } : {}),
            ...(anchor.fingerprint ? { fingerprint: anchor.fingerprint } : {}),
            inTable: anchor.inTable
        },
        modified: [anchor.exactText, ...missing].join('\n'),
        structuredContent: true
    };
    const result = await session.applyOperations([operation], {
        author,
        atomic: true,
        strictTargets: true,
        generateRedlines: false,
        sanitizeInput: false,
        structuredContent: true
    });
    if (result.status !== 'ok' || result.written !== true) {
        throw new Error(result.error?.message || 'Could not seed kitchen-sink target paragraphs.');
    }

    const after = session.inspect();
    if (after.status !== 'ok') throw new Error(after.error?.message || 'Could not inspect document after marker seeding.');
    const counts = new Map();
    for (const paragraph of after.paragraphs) {
        counts.set(paragraph.exactText, (counts.get(paragraph.exactText) || 0) + 1);
    }
    const missingAfter = missing.filter(marker => counts.get(marker) !== 1);
    const sourceCounts = new Map();
    const afterCounts = new Map();
    for (const paragraph of before.paragraphs) {
        sourceCounts.set(paragraph.exactText, (sourceCounts.get(paragraph.exactText) || 0) + 1);
    }
    for (const paragraph of after.paragraphs) {
        afterCounts.set(paragraph.exactText, (afterCounts.get(paragraph.exactText) || 0) + 1);
    }
    const missingSource = [...sourceCounts].filter(([text, count]) => afterCounts.get(text) !== count);
    const expectedParagraphCount = before.paragraphs.length + missing.length;
    if (missingAfter.length || missingSource.length || after.paragraphs.length !== expectedParagraphCount) {
        throw new Error('Structured marker seeding did not preserve every source paragraph and create one paragraph per marker.');
    }
    return { added: missing, inspection: after, result };
}

/**
 * Opens a browser DOCX editing session through the package's public document
 * facade. All mutation requests are submitted as one atomic operation batch;
 * the wrapper does not expose ZIP editing or Word host APIs.
 */
export function createDocumentSession(input) {
    const document = openDocx(input);

    const session = {
        inspect(options) {
            return document.inspect(options);
        },

        getPromptParagraphs() {
            const inspection = document.inspect();
            if (inspection.status !== 'ok') {
                throw new Error(inspection.error?.message || 'DOCX inspection failed');
            }

            const parsed = parseOoxmlSafe(requiredDocumentXml(document), 'application/xml');
            if (!parsed.doc || parsed.error) {
                throw new Error(parsed.error?.message || 'Could not parse word/document.xml');
            }
            const paragraphNodes = getDocumentParagraphNodes(parsed.doc);
            return inspection.paragraphs.flatMap((paragraph, index) => {
                const text = String(paragraph.exactText || '').trim();
                if (!text) return [];
                const node = paragraphNodes[index];
                const projectedText = node ? formatParagraphNode(node) : '';
                const formattedText = formatPromptParagraph(paragraph, projectedText || text);
                return [{
                    index: paragraph.index,
                    text,
                    formattedText
                }];
            });
        },

        async applyOperations(operations, options = {}) {
            if (!Array.isArray(operations)) {
                throw new TypeError('Document operations must be an array');
            }
            const { atomic: _ignoredAtomic, ...callerOptions } = options || {};
            return document.applyOperations(operations, {
                ...callerOptions,
                atomic: true,
                strictTargets: callerOptions.strictTargets !== false
            });
        },

        getDocumentXml() {
            return requiredDocumentXml(document);
        },

        toUint8Array() {
            return document.toUint8Array();
        }
    };

    return Object.freeze(session);
}
