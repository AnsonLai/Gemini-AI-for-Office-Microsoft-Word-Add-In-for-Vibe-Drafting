import { configureXmlProvider } from '../src/taskpane/modules/reconciliation/adapters/xml-adapter.js';

let DOMParserCtor = null;
let XMLSerializerCtor = null;
let doc = null;
let win = null;
let nodeCtor = null;

function ensureIterable(collection) {
    if (!collection) return;
    const proto = Object.getPrototypeOf(collection);
    if (!proto || proto[Symbol.iterator]) return;
    Object.defineProperty(proto, Symbol.iterator, {
        configurable: true,
        writable: true,
        value: function* iterator() {
            for (let i = 0; i < this.length; i++) {
                if (typeof this.item === 'function') {
                    yield this.item(i);
                } else {
                    yield this[i];
                }
            }
        }
    });
}

function ensureRemove(node) {
    if (!node) return;
    const proto = Object.getPrototypeOf(node);
    if (!proto || proto.remove) return;
    Object.defineProperty(proto, 'remove', {
        configurable: true,
        writable: true,
        value: function remove() {
            if (this.parentNode) {
                this.parentNode.removeChild(this);
            }
        }
    });
}

try {
    const { JSDOM } = await import('jsdom');
    const dom = new JSDOM('');
    DOMParserCtor = dom.window.DOMParser;
    XMLSerializerCtor = dom.window.XMLSerializer;
    doc = dom.window.document;
    win = dom.window;
    nodeCtor = dom.window.Node;
} catch {
    const xmldom = await import('@xmldom/xmldom');
    DOMParserCtor = xmldom.DOMParser;
    XMLSerializerCtor = xmldom.XMLSerializer;
    const parser = new DOMParserCtor();
    doc = parser.parseFromString('<root/>', 'text/xml');
    ensureIterable(doc.getElementsByTagName('*'));
    ensureIterable(doc.childNodes);
    ensureRemove(doc.documentElement);
    ensureRemove(doc.documentElement?.firstChild);
    win = {
        DOMParser: DOMParserCtor,
        XMLSerializer: XMLSerializerCtor,
        document: doc
    };
    nodeCtor = {
        ELEMENT_NODE: 1,
        ATTRIBUTE_NODE: 2,
        TEXT_NODE: 3,
        CDATA_SECTION_NODE: 4,
        ENTITY_REFERENCE_NODE: 5,
        ENTITY_NODE: 6,
        PROCESSING_INSTRUCTION_NODE: 7,
        COMMENT_NODE: 8,
        DOCUMENT_NODE: 9,
        DOCUMENT_TYPE_NODE: 10,
        DOCUMENT_FRAGMENT_NODE: 11
    };
}

configureXmlProvider({
    DOMParser: DOMParserCtor,
    XMLSerializer: XMLSerializerCtor
});

if (!global.DOMParser) global.DOMParser = DOMParserCtor;
if (!global.XMLSerializer) global.XMLSerializer = XMLSerializerCtor;
if (!global.document) global.document = doc;
if (!global.window) global.window = win;
if (!global.Node) global.Node = nodeCtor;
