/**
 * Shared XML helpers for ingestion flows.
 */

/**
 * Returns all children as an array.
 *
 * @param {Node} node - Parent node
 * @returns {Node[]}
 */
export function childNodesToArray(node) {
    return Array.from(node?.childNodes || []);
}

/**
 * Serializes all attributes into a plain string.
 *
 * @param {Element} element - Element with attributes
 * @returns {string}
 */
export function serializeAttributes(element) {
    return Array.from(element.attributes)
        .map(attr => `${attr.name}="${attr.value}"`)
        .join(' ');
}

/**
 * Returns true when a node belongs to the target namespace and optional local name.
 *
 * @param {Node} node - Candidate node
 * @param {string} namespaceUri - Namespace URI
 * @param {string} [localName] - Optional local name
 * @returns {boolean}
 */
export function isNamespacedNode(node, namespaceUri, localName = '') {
    if (!node || node.namespaceURI !== namespaceUri) return false;
    if (!localName) return true;
    return node.localName === localName;
}
