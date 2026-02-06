/**
 * XML adapter for parser/serializer portability.
 *
 * Default behavior uses browser-provided DOMParser/XMLSerializer.
 * Consumers can override these constructors for non-browser runtimes.
 */

let _DOMParser = globalThis.DOMParser;
let _XMLSerializer = globalThis.XMLSerializer;

/**
 * Configures XML provider constructors.
 *
 * @param {Object} [options={}] - Provider overrides
 * @param {typeof DOMParser} [options.DOMParser] - DOMParser constructor
 * @param {typeof XMLSerializer} [options.XMLSerializer] - XMLSerializer constructor
 */
export function configureXmlProvider(options = {}) {
    if (options.DOMParser) _DOMParser = options.DOMParser;
    if (options.XMLSerializer) _XMLSerializer = options.XMLSerializer;
}

/**
 * Creates a parser instance.
 *
 * @returns {DOMParser}
 */
export function createParser() {
    if (!_DOMParser && globalThis.DOMParser) {
        _DOMParser = globalThis.DOMParser;
    }
    if (!_DOMParser) {
        throw new Error('DOMParser is not configured. Call configureXmlProvider({ DOMParser, XMLSerializer }) first.');
    }
    return new _DOMParser();
}

/**
 * Creates a serializer instance.
 *
 * @returns {XMLSerializer}
 */
export function createSerializer() {
    if (!_XMLSerializer && globalThis.XMLSerializer) {
        _XMLSerializer = globalThis.XMLSerializer;
    }
    if (!_XMLSerializer) {
        throw new Error('XMLSerializer is not configured. Call configureXmlProvider({ DOMParser, XMLSerializer }) first.');
    }
    return new _XMLSerializer();
}

/**
 * Parses XML text into a DOM document.
 *
 * @param {string} xmlString - XML string
 * @param {string} [contentType='text/xml'] - MIME type
 * @returns {Document}
 */
export function parseXml(xmlString, contentType = 'text/xml') {
    const parser = createParser();
    return parser.parseFromString(xmlString, contentType);
}

/**
 * Serializes a node to XML text.
 *
 * @param {Node} node - Node to serialize
 * @returns {string}
 */
export function serializeXml(node) {
    const serializer = createSerializer();
    return serializer.serializeToString(node);
}
