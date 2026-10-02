// Golden scenario: the owner's 16-prompt editing session on the sample NDA.
//
// Each step has:
//   prompt  - typed verbatim into the real taskpane chat (both lanes)
//   replay  - deterministic lane only: the model responses to serve, in order.
//             { chat: <Gemini part> } answers a chat-model call; { aux: <JSON> }
//             answers a tool's secondary JSON call (redline diff, comments,
//             highlights). Paragraph references are resolved at runtime from
//             the [P#] context inside the intercepted request:
//               { $p: <matcher>, nth? }   -> paragraph number
//               { $all: <matcher> }       -> array of paragraph numbers
//               { $anchor: <matcher> }    -> first 40 chars of that paragraph
//   checks  - scored on the accepted view of the exported document (both lanes).
//             Earlier checks are re-run at later steps unless persist: false.
//   knownIssue - links a documented upstream library defect expected to fail.
//
// Matchers: "substring" | { equals } | { startsWith } | { includes } | { regex, flags }.

const done = { chat: { text: 'Done.' } };
const call = (name, args) => ({ chat: { functionCall: { name, args } } });
const redline = instruction => call('apply_redlines', { instruction });
const p = (matcher, nth) => ({ $p: matcher, ...(nth ? { nth } : {}) });
const anchor = matcher => ({ $anchor: matcher });

const HEADERS = [
  'DEFINITION OF CONFIDENTIAL INFORMATION', 'EXCLUSIONS', 'OBLIGATIONS OF RECEIVING PARTY', 'TERM',
  'REQUIRED DISCLOSURE', 'RETURN OF INFORMATION', 'REMEDIES', 'GOVERNING LAW', 'GENERAL PROVISIONS'
];
const RECITALS = [
  'The Disclosing Party possesses certain confidential, proprietary, and trade secret information.',
  'The Parties desire to enter into a potential business relationship or transaction (the “Purpose”), which requires the Disclosing Party to disclose certain Confidential Information (as defined below) to the Receiving Party.',
  'The Receiving Party agrees to receive and treat such Confidential Information in confidence, subject to the terms and conditions of this Agreement.'
];
const GOVERNING_LAW = { includes: 'governed by and construed' };

export const scenario = {
  name: 'nda-golden',
  fixture: 'tests/fixtures/golden/sample-nda.docx',
  steps: [
    {
      id: '01-instructions',
      prompt: 'can you add a paragraph at the top before the title as instructions for the user?  start with a bolded "INSTRUCTIONS:" so users know what it is',
      replay: [
        redline('Insert a new first paragraph before the title, starting with a bold "INSTRUCTIONS:" label, explaining how to complete the agreement.'),
        { aux: [{
          paragraphIndex: p({ equals: 'NON-DISCLOSURE AGREEMENT' }),
          anchorText: anchor({ equals: 'NON-DISCLOSURE AGREEMENT' }),
          operation: 'edit_paragraph',
          newContent: '**INSTRUCTIONS:** Complete the bracketed fields, review each section with counsel, and have both parties sign and date the signature block below.\nNON-DISCLOSURE AGREEMENT'
        }] },
        done
      ],
      checks: [
        { type: 'paragraph', label: 'instructions are the first paragraph', index: 1, text: { startsWith: 'INSTRUCTIONS:' } },
        { type: 'format', label: '"INSTRUCTIONS:" is bold', find: 'INSTRUCTIONS:', paragraph: { startsWith: 'INSTRUCTIONS:' }, occurrence: 1, expect: { bold: true } },
        { type: 'adjacent', label: 'title follows the instructions', first: { startsWith: 'INSTRUCTIONS:' }, second: { equals: 'NON-DISCLOSURE AGREEMENT' } }
      ]
    },
    {
      id: '02-underline-title',
      prompt: 'can you underline the title?',
      replay: [
        redline('Underline the title "NON-DISCLOSURE AGREEMENT".'),
        { aux: [{
          paragraphIndex: p({ equals: 'NON-DISCLOSURE AGREEMENT' }),
          anchorText: anchor({ equals: 'NON-DISCLOSURE AGREEMENT' }),
          operation: 'format_text',
          find: 'NON-DISCLOSURE AGREEMENT',
          formatting: { underline: true }
        }] },
        done
      ],
      checks: [
        { type: 'format', label: 'title is underlined', find: 'NON-DISCLOSURE AGREEMENT', paragraph: { equals: 'NON-DISCLOSURE AGREEMENT' }, all: true, expect: { underline: true } },
        { type: 'format', label: 'title keeps its bold', find: 'NON-DISCLOSURE AGREEMENT', paragraph: { equals: 'NON-DISCLOSURE AGREEMENT' }, all: true, expect: { bold: true } }
      ]
    },
    {
      id: '03-recitals-list',
      prompt: 'can you turn the recitals into a proper ordered list?  A, B, C...',
      replay: [
        call('edit_list', {
          startParagraphIndex: p({ startsWith: 'A. The Disclosing Party possesses' }),
          endParagraphIndex: p({ startsWith: 'A. The Disclosing Party possesses' }),
          newItems: RECITALS,
          listType: 'numbered',
          numberingStyle: 'upperAlpha'
        }),
        done
      ],
      checks: [
        { type: 'list', label: 'recitals are an A/B/C list', items: RECITALS.map(text => ({ startsWith: text.slice(0, 30) })), numFmt: 'upperLetter' },
        { type: 'text', label: 'manual "A." marker removed', text: { regex: '(^|\\n)A\\. The Disclosing Party possesses' }, absent: true }
      ]
    },
    {
      id: '04-parties-table',
      prompt: 'can you turn the parties at the top into a table to save space?',
      replay: [
        redline('Replace the Disclosing Party, "And", and Receiving Party paragraphs with a two-column table: one column per party, rows for name, address, and defined term.'),
        { aux: [{
          paragraphIndex: p({ startsWith: 'Disclosing Party:' }),
          endParagraphIndex: p({ startsWith: 'Receiving Party:' }),
          anchorText: anchor({ startsWith: 'Disclosing Party:' }),
          operation: 'replace_range',
          content: '| Disclosing Party | Receiving Party |\n|---|---|\n| [Name of Disclosing Party] | [Name of Receiving Party] |\n| [Address of Disclosing Party] | [Address of Receiving Party] |\n| (the “Disclosing Party”) | (the “Receiving Party”) |'
        }] },
        done
      ],
      checks: [
        { type: 'tableRow', label: 'party names share a table row', row: ['[Name of Disclosing Party]', '[Name of Receiving Party]'] },
        { type: 'tableRow', label: 'party addresses share a table row', row: ['[Address of Disclosing Party]', '[Address of Receiving Party]'] },
        { type: 'paragraph', label: 'standalone "And" removed', text: { equals: 'And' }, count: 0 }
      ]
    },
    {
      id: '05-bold-signature-labels',
      prompt: 'can you bold the By and Title in the signature block?',
      replay: [
        redline('In the signature block table, bold the "By:" and "Title:" labels in both columns.'),
        { aux: [
          { paragraphIndex: p({ startsWith: 'By: [Name]' }, 1), anchorText: 'By: [Name]', operation: 'format_text', find: 'By:', formatting: { bold: true } },
          { paragraphIndex: p({ startsWith: 'By: [Name]' }, 2), anchorText: 'By: [Name]', operation: 'format_text', find: 'By:', formatting: { bold: true } },
          { paragraphIndex: p({ equals: 'Title:' }, 1), anchorText: 'Title:', operation: 'format_text', find: 'Title:', formatting: { bold: true } },
          { paragraphIndex: p({ equals: 'Title:' }, 2), anchorText: 'Title:', operation: 'format_text', find: 'Title:', formatting: { bold: true } }
        ] },
        done
      ],
      checks: [
        { type: 'format', label: 'every "By:" is bold', find: 'By:', paragraph: { startsWith: 'By:' }, inTable: true, all: true, expect: { bold: true } },
        { type: 'format', label: 'every "Title:" is bold', find: 'Title:', paragraph: { startsWith: 'Title:' }, inTable: true, all: true, expect: { bold: true } }
      ]
    },
    {
      id: '06-signature-date-row',
      prompt: 'can you add another row to the signature block table for dates?',
      replay: [
        call('edit_table', {
          paragraphIndex: p({ startsWith: 'By: [Name]' }),
          action: 'add_row',
          content: [['Date: _______________', 'Date: _______________']]
        }),
        done
      ],
      checks: [
        { type: 'tableRow', label: 'signature table has a Date row', tableContaining: { startsWith: 'By:' }, row: [{ startsWith: 'Date' }, { startsWith: 'Date' }] },
        { type: 'tableShape', label: 'signature table grew to 5 rows', tableContaining: { startsWith: 'By:' }, rows: 5 }
      ]
    },
    {
      id: '07-confidential-bullet',
      prompt: 'can you add a new bullet for the examples of confidential information between 2 and 3 that includes photographs, videos, and other recordings of prototypes and physical hardware?',
      replay: [
        call('insert_list_item', {
          afterParagraphIndex: p({ startsWith: 'Technical data, specifications' }),
          text: 'Photographs, videos, and other recordings of prototypes and physical hardware.'
        }),
        done
      ],
      checks: [
        { type: 'list', label: 'new example sits between items 2 and 3', items: [
          { startsWith: 'Business plans' }, { startsWith: 'Technical data' },
          { regex: 'photographs.*videos.*recordings.*prototypes.*hardware' }, { startsWith: 'Information concerning' }
        ] }
      ]
    },
    {
      id: '08-archival-two-copies',
      prompt: 'can you change it so that the archival exception allows for 2 copies?',
      replay: [
        redline('In the Archival Exception, allow legal counsel to retain two copies instead of one.'),
        { aux: [{
          paragraphIndex: p({ includes: 'legal counsel may retain' }),
          anchorText: anchor({ includes: 'legal counsel may retain' }),
          operation: 'edit_paragraph',
          replacements: [{ find: 'one (1) copy', replace: 'two (2) copies' }]
        }] },
        done
      ],
      checks: [
        { type: 'text', label: 'archival exception allows two copies', text: { regex: 'retain (two \\(2\\)|2|two) cop(y|ies)' } },
        { type: 'text', label: 'single-copy wording removed', text: 'one (1) copy', absent: true }
      ]
    },
    {
      id: '09-archival-subbullet',
      prompt: 'can you add a subbullet in the archival exception after 2.2 as 2.2.1 to clarify that the copy for archival purposes must be legally required by the SEC or FCC specifically?',
      replay: [
        call('insert_list_item', {
          afterParagraphIndex: p({ startsWith: 'This copy is to be used solely' }),
          text: 'Any copy retained for archival purposes must be legally required by the SEC or FCC specifically.',
          indentLevel: 1
        }),
        done
      ],
      checks: [
        { type: 'listItem', label: 'SEC/FCC clarification is a level-3 (2.2.1) item', text: { regex: 'SEC.*FCC|FCC.*SEC' }, level: 2 },
        { type: 'adjacent', label: 'clarification follows 2.2', first: { startsWith: 'This copy is to be used solely' }, second: { regex: 'SEC.*FCC|FCC.*SEC' } }
      ]
    },
    {
      id: '10-unbold-bc',
      prompt: 'can you unbold BC in governing law?',
      replay: [
        redline('In the governing law section, remove the bold from "British Columbia".'),
        { aux: [{
          paragraphIndex: p(GOVERNING_LAW),
          anchorText: anchor(GOVERNING_LAW),
          operation: 'format_text',
          find: 'British Columbia',
          occurrence: 1,
          formatting: { bold: false }
        }] },
        done
      ],
      checks: [
        { type: 'format', label: 'no "British Columbia" is bold', find: 'British Columbia', paragraph: GOVERNING_LAW, all: true, expect: { bold: false } },
        { type: 'text', label: 'governing law text unchanged', text: { includes: 'the laws of British Columbia, without regard' } }
      ]
    },
    {
      id: '11-required-disclosure-rewrite',
      prompt: 'can you rewrite the entire required disclosure provision so that up front, the receiving party should seek to fight any orders, will involve the disclosing party where possible, and only after those avenues are exhausted are they permitted to disclose.  Use a few bullets after the initial paragraph so it\'s easier to read.',
      replay: [
        redline('Rewrite the Required Disclosure provision body as an introductory paragraph followed by bullets: resist the order first, involve the Disclosing Party where possible, and disclose only after those avenues are exhausted.'),
        { aux: [{
          paragraphIndex: p({ startsWith: 'If the Receiving Party is required by law' }),
          anchorText: anchor({ startsWith: 'If the Receiving Party is required by law' }),
          operation: 'replace_paragraph',
          content: 'If the Receiving Party is required by law, regulation, or court order to disclose any Confidential Information, the Receiving Party shall first resist such disclosure and may disclose only as follows:\n- The Receiving Party shall use all reasonable legal means to contest, narrow, or quash the order or requirement.\n- Where legally permitted, the Receiving Party shall promptly notify the Disclosing Party in writing and involve it in seeking a protective order or other appropriate remedy.\n- Only after these avenues are exhausted may the Receiving Party disclose, and then only the portion of Confidential Information it is legally required to disclose.'
        }] },
        done
      ],
      checks: [
        { type: 'section', label: 'intro paragraph then bullets', header: { regex: 'REQUIRED DISCLOSURE' }, until: { regex: 'RETURN OF INFORMATION' },
          introParagraph: true, minBullets: 2, mentions: ['contest|resist|fight|challeng|oppos|quash', 'Disclosing Party', 'exhaust'] }
      ]
    },
    {
      id: '12-headers-to-list',
      prompt: 'can you turn all the headers (start with 1., 2., etc.) into a proper ordered list instead of just regular text?',
      replay: [
        call('convert_headers_to_list', {
          paragraphIndices: { $all: { regex: '^[1-9]\\. [A-Z][A-Z ]+$', flags: '' } },
          numberingFormat: 'arabic'
        }),
        done
      ],
      checks: [
        { type: 'sameList', label: 'all nine headers share one decimal list', texts: HEADERS.map(text => ({ equals: text })), numFmt: 'decimal' },
        { type: 'text', label: 'manual header numbers removed', text: { regex: '(^|\\n)[1-9]\\. (DEFINITION|EXCLUSIONS|GOVERNING LAW)' }, absent: true }
      ]
    },
    {
      id: '13-highlight-date',
      prompt: 'can you highlight the date of the agreement?',
      replay: [
        call('highlight_text', { instruction: 'Highlight the date of the agreement (the Effective Date).', color: 'yellow' }),
        { aux: [{ paragraphIndex: p({ startsWith: 'This Non-Disclosure Agreement' }), textToFind: 'the date of last signature below' }] },
        done
      ],
      checks: [
        { type: 'highlight', label: 'agreement date is highlighted', text: { regex: 'date|effective' } }
      ]
    },
    {
      id: '14-review-comments',
      prompt: 'can you add comments on where the document should be changed, just 1-2 of the biggest legal shortcomings?',
      replay: [
        call('insert_comment', { instruction: 'Add one or two comments on the biggest legal shortcomings in this NDA.' }),
        { aux: [
          { paragraphIndex: p({ startsWith: 'This Agreement shall commence on the Effective Date' }), textToFind: 'five (5) years',
            commentContent: 'Trade secrets lose protection when confidentiality expires. Consider making obligations for trade secrets survive for as long as they remain trade secrets.' },
          { paragraphIndex: p({ startsWith: 'The Receiving Party acknowledges that unauthorized disclosure' }), textToFind: 'irreparable harm',
            commentContent: 'Acknowledging irreparable harm is helpful, but the clause should expressly entitle the Disclosing Party to injunctive relief without posting a bond.' }
        ] },
        done
      ],
      checks: [
        { type: 'comments', label: 'one or two review comments', min: 1, max: 2 }
      ]
    },
    {
      id: '15-bold-purpose-phrase',
      prompt: 'can you bold "potential business relationship or transaction"?',
      replay: [
        redline('Bold the phrase "potential business relationship or transaction".'),
        { aux: [{
          paragraphIndex: p({ includes: 'potential business relationship or transaction' }),
          anchorText: anchor({ includes: 'potential business relationship or transaction' }),
          operation: 'format_text',
          find: 'potential business relationship or transaction',
          formatting: { bold: true }
        }] },
        done
      ],
      // Step 16 rewrites this phrase, so its checks are not regressed later.
      persist: false,
      checks: [
        { type: 'format', label: 'purpose phrase is bold', find: 'potential business relationship or transaction', all: true, expect: { bold: true } }
      ]
    },
    {
      id: '16-titan-rewrite',
      prompt: 'can you rewrite "potential business relationship or transaction" to be a specific project, codenamed Titan?',
      replay: [
        redline('Replace "a potential business relationship or transaction" with a specific project codenamed Titan.'),
        { aux: [{
          paragraphIndex: p({ includes: 'potential business relationship or transaction' }),
          anchorText: anchor({ includes: 'potential business relationship or transaction' }),
          operation: 'edit_paragraph',
          replacements: [{ find: 'a potential business relationship or transaction', replace: 'a specific project codenamed “Titan”' }]
        }] },
        done
      ],
      checks: [
        { type: 'text', label: 'purpose names project Titan', text: { regex: 'project.{0,40}Titan|Titan.{0,40}project' } },
        { type: 'text', label: 'generic purpose wording removed', text: 'potential business relationship or transaction', absent: true }
      ]
    }
  ]
};
