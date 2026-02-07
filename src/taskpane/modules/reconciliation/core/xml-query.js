/**
 * Shared XML query helpers for OOXML documents.
 *
 * These wrappers provide consistent first/all element access across:
 * - Namespace-aware lookups (`getElementsByTagNameNS`)
 * - Prefix-based fallbacks (`getElementsByTagName('w:...')`)
 */

/**
 * Returns all descendant elements matching a qualified tag name.
 *
 * @param {Node|Document|Element|null|undefined} node - Query root
 * @param {string} tagName - Qualified element name (for example, `w:p`)
 * @returns {Element[]}
 */
export function getElementsByTag(node, tagName) {
    if (!node || typeof node.getElementsByTagName !== 'function') return [];
    return Array.from(node.getElementsByTagName(tagName));
}

/**
 * Returns the first descendant element matching a qualified tag name.
 *
 * @param {Node|Document|Element|null|undefined} node - Query root
 * @param {string} tagName - Qualified element name (for example, `w:p`)
 * @returns {Element|null}
 */
export function getFirstElementByTag(node, tagName) {
    if (!node || typeof node.getElementsByTagName !== 'function') return null;
    const elements = node.getElementsByTagName(tagName);
    return elements.length > 0 ? elements[0] : null;
}

/**
 * Returns all descendant elements matching namespace + local name.
 *
 * @param {Node|Document|Element|null|undefined} node - Query root
 * @param {string} namespaceUri - Namespace URI (or `*`)
 * @param {string} localName - Local tag name (for example, `p`)
 * @returns {Element[]}
 */
export function getElementsByTagNS(node, namespaceUri, localName) {
    if (!node || typeof node.getElementsByTagNameNS !== 'function') return [];
    return Array.from(node.getElementsByTagNameNS(namespaceUri, localName));
}

/**
 * Returns the first descendant element matching namespace + local name.
 *
 * @param {Node|Document|Element|null|undefined} node - Query root
 * @param {string} namespaceUri - Namespace URI (or `*`)
 * @param {string} localName - Local tag name (for example, `p`)
 * @returns {Element|null}
 */
export function getFirstElementByTagNS(node, namespaceUri, localName) {
    if (!node || typeof node.getElementsByTagNameNS !== 'function') return null;
    const elements = node.getElementsByTagNameNS(namespaceUri, localName);
    return elements.length > 0 ? elements[0] : null;
}

/**
 * Returns all elements using namespace-aware lookup with prefixed fallback.
 *
 * @param {Node|Document|Element|null|undefined} node - Query root
 * @param {string} namespaceUri - Namespace URI
 * @param {string} localName - Local name
 * @param {string} [fallbackTagName] - Optional prefixed fallback (default: `w:${localName}`)
 * @returns {Element[]}
 */
export function getElementsByTagNSOrTag(node, namespaceUri, localName, fallbackTagName = `w:${localName}`) {
    const namespacedElements = getElementsByTagNS(node, namespaceUri, localName);
    if (namespacedElements.length > 0) return namespacedElements;
    return getElementsByTag(node, fallbackTagName);
}

/**
 * Returns the first element using namespace-aware lookup with prefixed fallback.
 *
 * @param {Node|Document|Element|null|undefined} node - Query root
 * @param {string} namespaceUri - Namespace URI
 * @param {string} localName - Local name
 * @param {string} [fallbackTagName] - Optional prefixed fallback (default: `w:${localName}`)
 * @returns {Element|null}
 */
export function getFirstElementByTagNSOrTag(node, namespaceUri, localName, fallbackTagName = `w:${localName}`) {
    const namespacedElement = getFirstElementByTagNS(node, namespaceUri, localName);
    if (namespacedElement) return namespacedElement;
    return getFirstElementByTag(node, fallbackTagName);
}

/**
 * Returns XML parser error element if present.
 *
 * @param {Document|Element|null|undefined} xmlDoc - Parsed XML document
 * @returns {Element|null}
 */
export function getXmlParseError(xmlDoc) {
    return getFirstElementByTag(xmlDoc, 'parsererror');
}
