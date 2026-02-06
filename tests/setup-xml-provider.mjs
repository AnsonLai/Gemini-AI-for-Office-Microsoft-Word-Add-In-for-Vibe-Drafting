import { configureXmlProvider } from '../src/taskpane/modules/reconciliation/adapters/xml-adapter.js';

let DOMParserCtor = null;
let XMLSerializerCtor = null;

try {
    const { JSDOM } = await import('jsdom');
    const dom = new JSDOM('');
    DOMParserCtor = dom.window.DOMParser;
    XMLSerializerCtor = dom.window.XMLSerializer;
} catch {
    const xmldom = await import('@xmldom/xmldom');
    DOMParserCtor = xmldom.DOMParser;
    XMLSerializerCtor = xmldom.XMLSerializer;
}

configureXmlProvider({
    DOMParser: DOMParserCtor,
    XMLSerializer: XMLSerializerCtor
});
