import crypto from 'node:crypto';

export class DocxSessionStore {
    constructor() {
        /** @type {Map<string, any>} */
        this.sessions = new Map();
    }

    /**
     * @param {{
     *   zip: any,
     *   documentXml: string,
     *   sourcePath?: string|null,
     *   defaultGenerateRedlines?: boolean
     * }} input
     */
    create(input) {
        const sessionId = crypto.randomUUID();
        const now = new Date().toISOString();

        const session = {
            sessionId,
            zip: input.zip,
            documentXml: input.documentXml,
            sourcePath: input.sourcePath ?? null,
            defaultGenerateRedlines: input.defaultGenerateRedlines ?? true,
            dirty: false,
            createdAt: now,
            updatedAt: now
        };

        this.sessions.set(sessionId, session);
        return session;
    }

    /**
     * @param {string} sessionId
     */
    get(sessionId) {
        const session = this.sessions.get(sessionId);
        if (!session) {
            throw new Error(`Session not found: ${sessionId}`);
        }
        return session;
    }

    /**
     * @param {string} sessionId
     * @returns {boolean}
     */
    close(sessionId) {
        return this.sessions.delete(sessionId);
    }

    /**
     * @param {any} session
     */
    touch(session) {
        session.updatedAt = new Date().toISOString();
    }
}

