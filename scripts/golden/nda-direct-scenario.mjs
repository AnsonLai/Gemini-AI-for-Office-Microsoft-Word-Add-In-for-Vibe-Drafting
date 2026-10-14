// Direct-tool replay of the owner's original 16-prompt NDA session.
// Prompts are sourced verbatim from the frozen legacy scenario; only supported
// edits are exercised through opaque inspect_document handles. Legacy outcomes
// for gated operations remain visible as deferred checks rather than passes.

import { scenario as legacyScenario } from './nda-scenario.mjs';

const inspect = (search, options = {}) => ({
  chat: { functionCall: { name: 'inspect_document', args: { search, limit: 8, ...options } } }
});
const apply = edits => ({ chat: { functionCall: { name: 'apply_document_edits', args: { edits } } } });
const handle = (matcher, nth) => ({ $handle: matcher, ...(nth ? { nth } : {}) });
const done = { chat: { text: 'Done.' } };
// The nine typed numbered headers of step 12, in document order.
const NUMBERED_HEADERS = [
  'DEFINITION OF CONFIDENTIAL INFORMATION', 'EXCLUSIONS', 'OBLIGATIONS OF RECEIVING PARTY', 'TERM',
  'REQUIRED DISCLOSURE', 'RETURN OF INFORMATION', 'REMEDIES', 'GOVERNING LAW', 'GENERAL PROVISIONS'
];
const headerPattern = (number, name) => ({ regex: `^${number}\\. ${name}$`, flags: '' });
const batchEvidence = {
  type: 'directToolEvidence',
  label: 'edit batches have one verified commit and no nested model calls',
  persist: false
};

const plans = [
  {
    replay: [
      inspect('NON-DISCLOSURE AGREEMENT'),
      apply([{ type: 'insert_paragraph', target: handle({ equals: 'NON-DISCLOSURE AGREEMENT' }), position: 'before',
        text: 'INSTRUCTIONS: Complete the bracketed fields, review each section with counsel, and have both parties sign and date the signature block below.' }]),
      inspect('INSTRUCTIONS:'),
      apply([{ type: 'format_text', target: handle({ startsWith: 'INSTRUCTIONS:' }), find: 'INSTRUCTIONS:', properties: { bold: true } }]),
      done
    ]
  },
  {
    replay: [
      inspect('NON-DISCLOSURE AGREEMENT'),
      apply([{ type: 'format_text', target: handle({ equals: 'NON-DISCLOSURE AGREEMENT' }), find: 'NON-DISCLOSURE AGREEMENT', properties: { underline: true } }]),
      done
    ]
  },
  {
    scopeDisposition: {
      status: 'unsupported',
      reason: 'Conversion of existing list structure is outside the retained direct-edit scope.',
      refusalPattern: '(?:cannot|can’t|can\'t|unsupported|outside).*(?:convert|list|numbering)'
    },
    refusal: 'I can’t convert existing list structure in this editing workflow. List conversion is unsupported, so I left the document unchanged.'
  },
  {
    scopeDisposition: {
      status: 'unsupported',
      reason: 'Creating a table from paragraphs is outside the retained direct-edit scope.',
      refusalPattern: '(?:cannot|can’t|can\'t|unsupported|outside).*(?:table|convert|create)'
    },
    refusal: 'I can’t convert those paragraphs into a table in this editing workflow. Table creation is unsupported, so I left the document unchanged.'
  },
  {
    replay: [
      inspect('Title:', { around: 2 }),
      apply([
        { type: 'format_text', target: handle({ startsWith: 'By: [Name]' }, 1), find: 'By:', properties: { bold: true } },
        { type: 'format_text', target: handle({ startsWith: 'By: [Name]' }, 2), find: 'By:', properties: { bold: true } },
        { type: 'format_text', target: handle({ equals: 'Title:' }, 1), find: 'Title:', properties: { bold: true } },
        { type: 'format_text', target: handle({ equals: 'Title:' }, 2), find: 'Title:', properties: { bold: true } }
      ]),
      done
    ]
  },
  {
    scopeDisposition: {
      status: 'unsupported',
      reason: 'Adding a table row is outside the retained direct-edit scope.',
      refusalPattern: '(?:cannot|can’t|can\'t|unsupported|outside).*(?:table|row|add)'
    },
    refusal: 'I can’t add a row to that table in this editing workflow. Table row changes are unsupported, so I left the document unchanged.'
  },
  {
    replay: [
      inspect('Technical data, specifications'),
      apply([{ type: 'insert_list_item', target: handle({ startsWith: 'Technical data, specifications' }),
        text: 'Photographs, videos, and other recordings of prototypes and physical hardware.' }]),
      done
    ]
  },
  {
    replay: [
      inspect('one (1) copy'),
      apply([{ type: 'replace_text', target: handle({ includes: 'one (1) copy' }), find: 'one (1) copy', replace: 'two (2) copies' }]),
      done
    ]
  },
  {
    scopeDisposition: {
      status: 'unsupported',
      reason: 'Deep list-level changes are outside the retained direct-edit scope.',
      refusalPattern: '(?:cannot|can’t|can\'t|unsupported|outside).*(?:deep|nested|list|level|subbullet)'
    },
    refusal: 'I can’t add a nested subbullet in this editing workflow. Deep list-level changes are unsupported, so I left the document unchanged.'
  },
  {
    replay: [
      inspect('British Columbia', { around: 1 }),
      apply([{ type: 'format_text', target: handle({ includes: 'governed by and construed' }), find: 'British Columbia', occurrence: 1,
        properties: { bold: false } }]),
      done
    ]
  },
  {
    scopeDisposition: {
      status: 'unsupported',
      reason: 'Restructuring one provision into a paragraph and bullets is outside the retained direct-edit scope.',
      refusalPattern: '(?:cannot|can’t|can\'t|unsupported|outside).*(?:bullets|restructure|structure|paragraph)'
    },
    refusal: 'I can’t restructure that provision into an introductory paragraph and bullets in this editing workflow. Structural reformatting is unsupported, so I left the document unchanged.'
  },
  {
    // One parallel inspection turn finds all nine headers, then one convert_to_list per header in one batch.
    // The library joins the consecutive values 1. to 9. into one list.
    replay: [
      { chat: NUMBERED_HEADERS.map((name, index) => inspect(`${index + 1}. ${name}`).chat) },
      apply(NUMBERED_HEADERS.map((name, index) => ({ type: 'convert_to_list', target: handle(headerPattern(index + 1, name)) }))),
      done
    ]
  },
  {
    replay: [
      inspect('date of last signature below'),
      apply([{ type: 'format_text', target: handle({ includes: 'date of last signature below' }), find: 'date of last signature below',
        properties: { highlight: 'yellow' } }]),
      done
    ]
  },
  {
    replay: [
      inspect('five (5) years'),
      apply([{ type: 'comment', target: handle({ includes: 'five (5) years' }), anchor: 'five (5) years',
        text: 'Consider making confidentiality obligations for trade secrets survive for as long as the information remains a trade secret.' }]),
      done
    ]
  },
  {
    replay: [
      inspect('potential business relationship or transaction'),
      apply([{ type: 'format_text', target: handle({ includes: 'potential business relationship or transaction' }),
        find: 'potential business relationship or transaction', properties: { bold: true } }]),
      done
    ]
  },
  {
    replay: [
      inspect('potential business relationship or transaction'),
      apply([{ type: 'replace_text', target: handle({ includes: 'potential business relationship or transaction' }),
        find: 'a potential business relationship or transaction', replace: 'a specific project codenamed Titan' }]),
      done
    ]
  }
];

export const scenario = {
  name: 'nda-direct',
  protocol: 'direct-document-tools',
  fixture: legacyScenario.fixture,
  steps: legacyScenario.steps.map((legacyStep, index) => {
    const plan = plans[index];
    const base = { id: legacyStep.id, prompt: legacyStep.prompt };
    if (plan.scopeDisposition) {
      return {
        ...base,
        scopeDisposition: plan.scopeDisposition,
        persist: false,
        replay: [{ chat: { text: plan.refusal } }],
        checks: [{ type: 'scopeDisposition', label: 'unsupported operation explicitly refused without changing the document' }],
        deferredChecks: legacyStep.checks
      };
    }
    const checks = [...legacyStep.checks, batchEvidence];
    if (legacyStep.id === '15-bold-purpose-phrase') {
      // Step 16 intentionally replaces this phrase in the same user sequence.
      const purposeCheck = checks.findIndex(check => check.label === 'purpose phrase is bold');
      if (purposeCheck >= 0) checks[purposeCheck] = { ...checks[purposeCheck], persist: false };
    }
    return { ...base, replay: plan.replay, checks };
  })
};
