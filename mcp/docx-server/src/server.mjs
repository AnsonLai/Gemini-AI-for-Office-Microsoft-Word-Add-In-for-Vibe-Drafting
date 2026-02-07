#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import {
    createNewDocxPackage,
    ensureCommentsArtifacts,
    ensureNumberingArtifacts,
    loadDocxFromPath,
    normalizeDocumentXml,
    saveDocxSessionToPath
} from './services/docx-package-service.mjs';
import { DocxSessionStore } from './services/docx-session-store.mjs';
import {
    listParagraphs,
    replaceParagraph,
    resolveParagraph,
    serializeParagraph
} from './services/paragraph-targeting-service.mjs';
import {
    deriveParagraphAcceptedText,
    reconcileAddComment,
    reconcileParagraphEdit
} from './services/reconciliation-service.mjs';

const server = new Server(
    {
        name: 'reconciliation-docx-local',
        version: '0.1.0'
    },
    {
        capabilities: {
            tools: {}
        }
    }
);

const sessions = new DocxSessionStore();

const tools = [
    {
        name: 'docx_new',
        description: 'Create a new minimal valid .docx session. Optionally save it immediately.',
        inputSchema: {
            type: 'object',
            properties: {
                outputPath: { type: 'string', description: 'Optional output path to write the new .docx immediately.' },
                title: { type: 'string', description: 'Optional initial first-paragraph text.' },
                generateRedlines: { type: 'boolean', description: 'Default redline behavior for this session (default: true).' }
            },
            additionalProperties: false
        }
    },
    {
        name: 'docx_open',
        description: 'Open an existing .docx file as an editable session.',
        inputSchema: {
            type: 'object',
            properties: {
                path: { type: 'string', description: 'Path to an existing .docx file.' },
                generateRedlines: { type: 'boolean', description: 'Default redline behavior for this session (default: true).' }
            },
            required: ['path'],
            additionalProperties: false
        }
    },
    {
        name: 'docx_list_paragraphs',
        description: 'List paragraph handles and text for a session.',
        inputSchema: {
            type: 'object',
            properties: {
                sessionId: { type: 'string' },
                start: { type: 'integer', minimum: 0 },
                limit: { type: 'integer', minimum: 1, maximum: 500 }
            },
            required: ['sessionId'],
            additionalProperties: false
        }
    },
    {
        name: 'docx_edit_paragraph',
        description: 'Edit one paragraph using the reconciliation engine.',
        inputSchema: {
            type: 'object',
            properties: {
                sessionId: { type: 'string' },
                paragraphId: { type: 'string', description: 'Use id returned by docx_list_paragraphs.' },
                newText: { type: 'string', description: 'Replacement text, supports markdown format hints.' },
                author: { type: 'string' },
                generateRedlines: { type: 'boolean', description: 'Per-call override for redlines on/off.' }
            },
            required: ['sessionId', 'paragraphId', 'newText'],
            additionalProperties: false
        }
    },
    {
        name: 'docx_add_comment',
        description: 'Add a Word comment anchored to text within a paragraph.',
        inputSchema: {
            type: 'object',
            properties: {
                sessionId: { type: 'string' },
                paragraphId: { type: 'string' },
                textToFind: { type: 'string' },
                comment: { type: 'string' },
                author: { type: 'string' }
            },
            required: ['sessionId', 'paragraphId', 'textToFind', 'comment'],
            additionalProperties: false
        }
    },
    {
        name: 'docx_save_as',
        description: 'Save a session to a .docx path.',
        inputSchema: {
            type: 'object',
            properties: {
                sessionId: { type: 'string' },
                outputPath: { type: 'string' }
            },
            required: ['sessionId', 'outputPath'],
            additionalProperties: false
        }
    },
    {
        name: 'docx_close',
        description: 'Close a session and release memory.',
        inputSchema: {
            type: 'object',
            properties: {
                sessionId: { type: 'string' }
            },
            required: ['sessionId'],
            additionalProperties: false
        }
    }
];

server.setRequestHandler(ListToolsRequestSchema, async () => {
    return { tools };
});

server.setRequestHandler(CallToolRequestSchema, async request => {
    const toolName = request.params.name;
    const args = request.params.arguments || {};

    try {
        switch (toolName) {
            case 'docx_new':
                return await runDocxNew(args);
            case 'docx_open':
                return await runDocxOpen(args);
            case 'docx_list_paragraphs':
                return runDocxListParagraphs(args);
            case 'docx_edit_paragraph':
                return await runDocxEditParagraph(args);
            case 'docx_add_comment':
                return await runDocxAddComment(args);
            case 'docx_save_as':
                return await runDocxSaveAs(args);
            case 'docx_close':
                return runDocxClose(args);
            default:
                throw new Error(`Unknown tool: ${toolName}`);
        }
    } catch (error) {
        return errorResult(error);
    }
});

await server.connect(new StdioServerTransport());

async function runDocxNew(args) {
    const defaultGenerateRedlines = resolveRedlineMode(args.generateRedlines, true);
    const created = await createNewDocxPackage({
        title: args.title || ''
    });

    const session = sessions.create({
        zip: created.zip,
        documentXml: created.documentXml,
        sourcePath: null,
        defaultGenerateRedlines
    });

    let saved = null;
    if (args.outputPath) {
        saved = await saveDocxSessionToPath(session, args.outputPath);
        session.sourcePath = saved.outputPath;
    }

    const preview = listParagraphs(session.documentXml, { start: 0, limit: 5 });
    return okResult({
        sessionId: session.sessionId,
        defaultGenerateRedlines: session.defaultGenerateRedlines,
        paragraphs: preview,
        saved
    });
}

async function runDocxOpen(args) {
    const loaded = await loadDocxFromPath(String(args.path));
    const defaultGenerateRedlines = resolveRedlineMode(args.generateRedlines, true);
    const session = sessions.create({
        zip: loaded.zip,
        documentXml: loaded.documentXml,
        sourcePath: loaded.sourcePath,
        defaultGenerateRedlines
    });

    const preview = listParagraphs(session.documentXml, { start: 0, limit: 5 });
    return okResult({
        sessionId: session.sessionId,
        sourcePath: session.sourcePath,
        defaultGenerateRedlines: session.defaultGenerateRedlines,
        paragraphs: preview
    });
}

function runDocxListParagraphs(args) {
    const session = sessions.get(String(args.sessionId));
    const listing = listParagraphs(session.documentXml, {
        start: args.start,
        limit: args.limit
    });

    return okResult({
        sessionId: session.sessionId,
        defaultGenerateRedlines: session.defaultGenerateRedlines,
        ...listing
    });
}

async function runDocxEditParagraph(args) {
    const session = sessions.get(String(args.sessionId));
    const resolved = resolveParagraph(session.documentXml, String(args.paragraphId));
    const paragraphXml = serializeParagraph(resolved.paragraph);
    const paragraphText = deriveParagraphAcceptedText(paragraphXml);
    const generateRedlines = resolveRedlineMode(args.generateRedlines, session.defaultGenerateRedlines);

    const recon = await reconcileParagraphEdit({
        paragraphXml,
        paragraphText,
        modifiedText: String(args.newText),
        paraId: resolved.paraId,
        author: args.author || 'MCP AI',
        generateRedlines
    });

    if (!recon.hasChanges) {
        return okResult({
            sessionId: session.sessionId,
            paragraphId: resolved.id,
            changed: false,
            generateRedlines
        });
    }

    let updatedXml = replaceParagraph(resolved.doc, resolved.paragraph, recon.replacementNodes);
    updatedXml = normalizeDocumentXml(updatedXml);

    session.documentXml = updatedXml;
    session.dirty = true;
    sessions.touch(session);

    if (recon.numberingXml) {
        await ensureNumberingArtifacts(session.zip, recon.numberingXml);
    }

    const updatedWindow = listParagraphs(session.documentXml, {
        start: Math.max(0, resolved.index - 1),
        limit: 1
    });
    const updatedItem = updatedWindow.items[0] || null;

    return okResult({
        sessionId: session.sessionId,
        paragraphId: updatedItem ? updatedItem.id : resolved.id,
        changed: true,
        generateRedlines,
        sourceType: recon.sourceType,
        updatedText: updatedItem ? updatedItem.text : ''
    });
}

async function runDocxAddComment(args) {
    const session = sessions.get(String(args.sessionId));
    const resolved = resolveParagraph(session.documentXml, String(args.paragraphId));

    const result = reconcileAddComment({
        documentXml: session.documentXml,
        paragraphIndex: resolved.index,
        textToFind: String(args.textToFind),
        commentContent: String(args.comment),
        author: args.author || 'MCP AI'
    });

    if (!result.commentsApplied) {
        return okResult({
            sessionId: session.sessionId,
            paragraphId: resolved.id,
            commentsApplied: 0,
            warnings: result.warnings || []
        });
    }

    session.documentXml = normalizeDocumentXml(result.oxml);
    session.dirty = true;
    sessions.touch(session);

    const mergeInfo = await ensureCommentsArtifacts(session.zip, result.commentsXml);

    return okResult({
        sessionId: session.sessionId,
        paragraphId: resolved.id,
        commentsApplied: result.commentsApplied,
        warnings: result.warnings || [],
        mergedComments: mergeInfo.addedComments
    });
}

async function runDocxSaveAs(args) {
    const session = sessions.get(String(args.sessionId));
    const saveResult = await saveDocxSessionToPath(session, String(args.outputPath));
    session.sourcePath = saveResult.outputPath;
    session.dirty = false;
    sessions.touch(session);

    return okResult({
        sessionId: session.sessionId,
        outputPath: saveResult.outputPath,
        bytes: saveResult.bytes,
        dirty: session.dirty
    });
}

function runDocxClose(args) {
    const sessionId = String(args.sessionId);
    const closed = sessions.close(sessionId);
    return okResult({ sessionId, closed });
}

function okResult(payload) {
    return {
        content: [
            {
                type: 'text',
                text: JSON.stringify(payload, null, 2)
            }
        ]
    };
}

function errorResult(error) {
    return {
        isError: true,
        content: [
            {
                type: 'text',
                text: JSON.stringify({
                    error: error?.message || String(error)
                }, null, 2)
            }
        ]
    };
}

function resolveRedlineMode(value, fallback) {
    if (typeof value === 'boolean') return value;
    return fallback;
}
