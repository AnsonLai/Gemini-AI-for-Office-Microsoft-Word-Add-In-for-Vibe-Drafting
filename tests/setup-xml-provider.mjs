import { configureXmlProvider } from '@ansonlai/docx-redline-js';

// Default lane: configure nothing. With no native DOMParser in Node, the library falls back to
// its own @xmldom/xmldom dependency. In Word it uses WebView2's native DOMParser.
// Browser-like diagnostic lane: DOCX_TEST_XML_PROVIDER=jsdom (see docs/plans/KNOWN_BUGS.md KB-8).
// The provider is chosen explicitly, never by whichever DOM package happens to be installed.
const XML_PROVIDER = (process.env.DOCX_TEST_XML_PROVIDER || 'xmldom').toLowerCase();
if (XML_PROVIDER !== 'xmldom' && XML_PROVIDER !== 'jsdom') {
    throw new Error(`Unknown DOCX_TEST_XML_PROVIDER "${XML_PROVIDER}"; use "xmldom" or "jsdom".`);
}

if (XML_PROVIDER === 'jsdom') {
    const { JSDOM } = await import('jsdom');
    const { window } = new JSDOM('');
    configureXmlProvider({ DOMParser: window.DOMParser, XMLSerializer: window.XMLSerializer });
}
