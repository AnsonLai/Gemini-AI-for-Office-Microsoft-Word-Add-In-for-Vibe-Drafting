#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { DocxSessionStore } from './services/docx-session-store.mjs';
import {
    addDocumentComment,
    applyDocumentOperations,
    createDocument,
    editDocumentParagraph,
    listDocumentParagraphs,
    openDocument,
    saveDocument
} from './services/docx-document-service.mjs';

const server = new Server(
    {
        name: 'docx-redline-js-local',
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
        description: 'Edit one paragraph using the DOCX document lifecycle.',
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
        name: 'docx_apply_operations',
        description: 'Apply a batch of docx-redline-js operations atomically to a session.',
        inputSchema: {
            type: 'object',
            properties: {
                sessionId: { type: 'string' },
                operations: { type: 'array', minItems: 1, items: { type: 'object' } },
                author: { type: 'string' },
                generateRedlines: { type: 'boolean' }
            },
            required: ['sessionId', 'operations'],
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
            case 'docx_apply_operations':
                return await runDocxApplyOperations(args);
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
    const doc = await createDocument(args.title || '');
    const session = sessions.create({ doc, sourcePath: null, defaultGenerateRedlines });
    let saved = null;
    if (args.outputPath) {
        saved = await saveDocument(doc, args.outputPath);
        session.sourcePath = saved.outputPath;
    }
    return okResult({
        sessionId: session.sessionId,
        defaultGenerateRedlines,
        paragraphs: listDocumentParagraphs(doc, { start: 0, limit: 5 }),
        saved
    });
}

async function runDocxOpen(args) {
    const loaded = await openDocument(String(args.path));
    const defaultGenerateRedlines = resolveRedlineMode(args.generateRedlines, true);
    const session = sessions.create({ ...loaded, defaultGenerateRedlines });
    return okResult({
        sessionId: session.sessionId,
        sourcePath: session.sourcePath,
        defaultGenerateRedlines,
        paragraphs: listDocumentParagraphs(session.doc, { start: 0, limit: 5 })
    });
}

function runDocxListParagraphs(args) {
    const session = sessions.get(String(args.sessionId));
    return okResult({
        sessionId: session.sessionId,
        defaultGenerateRedlines: session.defaultGenerateRedlines,
        ...listDocumentParagraphs(session.doc, { start: args.start, limit: args.limit })
    });
}

async function runDocxEditParagraph(args) {
    const session = sessions.get(String(args.sessionId));
    const generateRedlines = resolveRedlineMode(args.generateRedlines, session.defaultGenerateRedlines);
    const edited = await editDocumentParagraph(session.doc, String(args.paragraphId), String(args.newText), {
        author: args.author,
        generateRedlines
    });
    if (edited.changed) {
        session.dirty = true;
        sessions.touch(session);
    }
    return okResult({
        sessionId: session.sessionId,
        paragraphId: edited.paragraphId,
        changed: edited.changed,
        generateRedlines,
        sourceType: 'package',
        updatedText: edited.updatedText
    });
}

async function runDocxApplyOperations(args) {
    const session = sessions.get(String(args.sessionId));
    if (!Array.isArray(args.operations) || args.operations.length === 0) {
        throw new Error('operations must be a non-empty array');
    }
    const result = await applyDocumentOperations(session.doc, args.operations, {
        author: args.author || 'MCP AI',
        generateRedlines: resolveRedlineMode(args.generateRedlines, session.defaultGenerateRedlines)
    });
    if (result.written) {
        session.dirty = true;
        sessions.touch(session);
    }
    return okResult({
        sessionId: session.sessionId,
        changed: Boolean(result.written),
        status: result.status,
        completion: result.completion,
        results: result.results,
        receipts: result.receipts,
        artifactsChanged: result.artifactsChanged
    });
}

async function runDocxAddComment(args) {
    const session = sessions.get(String(args.sessionId));
    const added = await addDocumentComment(
        session.doc,
        String(args.paragraphId),
        String(args.textToFind),
        String(args.comment),
        { author: args.author }
    );
    if (added.result.written) {
        session.dirty = true;
        sessions.touch(session);
    }
    return okResult({
        sessionId: session.sessionId,
        paragraphId: added.paragraphId,
        commentsApplied: added.commentsApplied,
        warnings: added.result.warnings || [],
        mergedComments: added.commentsApplied
    });
}

async function runDocxSaveAs(args) {
    const session = sessions.get(String(args.sessionId));
    const saved = await saveDocument(session.doc, String(args.outputPath));
    session.sourcePath = saved.outputPath;
    session.dirty = false;
    sessions.touch(session);
    return okResult({ sessionId: session.sessionId, ...saved, dirty: false });
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
                    error: error?.message || String(error),
                    ...(error?.code ? { code: error.code } : {}),
                    ...(error?.warnings?.length ? { warnings: error.warnings } : {}),
                    ...(error?.details ? { details: error.details } : {})
                }, null, 2)
            }
        ]
    };
}

function resolveRedlineMode(value, fallback) {
    if (typeof value === 'boolean') return value;
    return fallback;
}
