// Second golden scenario: ten requests on a short, Word-authored privacy notice for a fictional
// clinic (tests/fixtures/golden/privacy-notice.docx, built by scripts/create-privacy-notice-fixture.ps1).
// It exists so live-model fixes are not tuned to the NDA alone. Each step has a natural prompt,
// a deterministic replay plan through inspect_document handles, and checks on the accepted view.

const inspect = (search, options = {}) => ({
  chat: { functionCall: { name: 'inspect_document', args: { search, limit: 8, ...options } } }
});
const apply = edits => ({ chat: { functionCall: { name: 'apply_document_edits', args: { edits } } } });
const handle = (matcher, nth) => ({ $handle: matcher, ...(nth ? { nth } : {}) });
const done = { chat: { text: 'Done.' } };
const batchEvidence = {
  type: 'directToolEvidence',
  label: 'edit batches have one verified commit and no nested model calls',
  persist: false
};
const EMAIL = 'privacy@maplestreetdental.example';
const SHARING = { startsWith: 'We share your information only' };
const USES = ['To provide and coordinate your care.', 'To process payments and insurance claims.', 'To send appointment reminders.'];
const ADDRESS = ['Maple Street Dental Clinic', '12 Maple Street', 'Springfield'];
const NEW_USES = ['To meet our legal and regulatory duties.', 'To improve our services.'];
const SECURITY = ['How We Keep Your Information Safe',
  'We store paper records in locked cabinets and electronic records on secure systems that only authorised staff can access.',
  'We keep your records for as long as the law requires and then destroy them securely.'];
const RIGHTS = ['Ask to see the information we hold about you.', 'Ask us to correct information that is wrong.', 'Withdraw your consent where we rely on it.'];

const steps = [
  {
    id: '01-last-updated',
    prompt: 'Please fill in the last updated date as October 30, 2026.',
    replay: [
      inspect('Last updated'),
      apply([{ type: 'replace_text', target: handle({ startsWith: 'Last updated:' }), find: '[Date]', replace: 'October 30, 2026' }]),
      done
    ],
    checks: [
      { type: 'text', label: 'last updated date is filled in', text: { equals: 'Last updated: October 30, 2026' } },
      { type: 'text', label: 'the [Date] placeholder is gone', text: { includes: '[Date]' }, absent: true }
    ]
  },
  {
    id: '02-bold-clinic-name',
    prompt: 'Bold the clinic name in the opening paragraph.',
    replay: [
      inspect('This notice explains'),
      apply([{ type: 'format_text', target: handle({ startsWith: 'This notice explains' }), find: 'Maple Street Dental Clinic', properties: { bold: true } }]),
      done
    ],
    checks: [
      { type: 'format', label: 'clinic name is bold in the opening paragraph', find: 'Maple Street Dental Clinic', paragraph: { startsWith: 'This notice explains' }, occurrence: 1, expect: { bold: true } }
    ]
  },
  {
    id: '03-newsletter-bullet',
    prompt: 'Under How We Use Your Information, add a bullet saying we send our newsletter only if you opt in.',
    replay: [
      inspect('To send appointment reminders.'),
      apply([{ type: 'insert_list_item', target: handle({ equals: 'To send appointment reminders.' }), text: 'To send our newsletter, only if you opt in.' }]),
      done
    ],
    checks: [
      // Step 12 adds two bullets mid-list later, so check membership and the newsletter's place, not one fixed run.
      { type: 'sameList', label: 'the newsletter bullet joins the uses list', texts: [...USES.map(text => ({ equals: text })), { includes: 'newsletter' }], numFmt: 'bullet' },
      { type: 'list', label: 'the newsletter bullet follows the reminders bullet', items: [USES[2], { includes: 'newsletter' }], numFmt: 'bullet' }
    ]
  },
  {
    id: '04-never-sell',
    prompt: 'In the sharing section, add that we never sell your personal information.',
    replay: [
      inspect('We share your information'),
      apply([{ type: 'rewrite_paragraph', target: handle(SHARING),
        text: 'We share your information only with other health professionals involved in your care, your insurer, and service providers who help us run the clinic. We never sell your personal information.' }]),
      done
    ],
    checks: [
      { type: 'text', label: 'the sharing section says the clinic never sells personal information', text: { regex: 'never sell', flags: 'i' } },
      { type: 'text', label: 'the existing sharing wording is kept', text: { includes: 'service providers who help us run the clinic' } },
      // The paragraph ends with a footnote reference (library issue #6, KB-15).
      { type: 'footnotes', label: 'the sharing footnote reference survives the rewrite', ids: ['1'] }
    ]
  },
  {
    id: '05-rights-comment',
    prompt: 'Add a comment on the Your Rights heading asking legal to confirm the list is complete.',
    replay: [
      inspect('Your Rights'),
      apply([{ type: 'comment', target: handle({ equals: 'Your Rights' }), text: 'Legal: please confirm this list of rights is complete.' }]),
      done
    ],
    checks: [{ type: 'comments', label: 'one review comment was added', min: 1, max: 2 }]
  },
  {
    id: '06-highlight-email',
    prompt: "Highlight the privacy officer's email address.",
    replay: [
      inspect(EMAIL),
      apply([{ type: 'format_text', target: handle({ equals: EMAIL }), find: EMAIL, properties: { highlight: 'yellow' } }]),
      done
    ],
    checks: [
      { type: 'highlight', label: 'the email address is highlighted', text: { includes: EMAIL } },
      { type: 'comments', label: 'no unrequested comments', min: 0, max: 0 }
    ]
  },
  {
    id: '07-rights-letters',
    prompt: 'Change the Your Rights list to use a, b, c instead of numbers.',
    replay: [
      inspect('Ask to see the information'),
      apply([{ type: 'change_list_format', target: handle({ equals: RIGHTS[0] }), format: 'lowerLetter' }]),
      done
    ],
    checks: [{ type: 'list', label: 'the rights list uses a, b, c', items: RIGHTS, numFmt: 'lowerLetter' }]
  },
  {
    id: '08-contact-bullets',
    prompt: 'Turn the contact table into a bulleted list instead.',
    scopeDisposition: {
      status: 'unsupported',
      reason: 'Restructuring a table into a list is outside the retained direct-edit scope.',
      refusalPattern: '(?:cannot|can’t|can\'t|unable|unsupported|outside|not supported).*(?:table|list|convert)'
    },
    refusal: 'I can’t convert the contact table into a list in this editing workflow; restructuring tables is unsupported, so I left the document unchanged.'
  },
  {
    id: '09-delete-acknowledgement',
    prompt: 'Delete the sentence saying the patient has read and understood the notice.',
    replay: [
      inspect('I have read and understood'),
      apply([{ type: 'delete_paragraph', target: handle({ startsWith: 'I have read and understood' }) }]),
      done
    ],
    checks: [
      { type: 'text', label: 'the acknowledgement sentence is gone', text: { includes: 'I have read and understood' }, absent: true },
      { type: 'text', label: 'the signature line is kept', text: { startsWith: 'Patient signature:' } }
    ]
  },
  {
    id: '10-paper-copy',
    prompt: 'Add a last line telling patients they can ask for a paper copy at the front desk.',
    replay: [
      inspect('Patient signature', { includeDocumentEnd: true }),
      apply([{ type: 'append_paragraph', target: { $end: true }, text: 'You can ask for a paper copy of this notice at the front desk.' }]),
      done
    ],
    checks: [
      { type: 'text', label: 'a paper-copy line is added', text: { regex: 'paper copy', flags: 'i' } },
      { type: 'adjacent', label: 'the paper-copy line follows the signature line', first: { startsWith: 'Patient signature:' }, second: { regex: 'paper copy', flags: 'i' } }
    ]
  },
  {
    // Several lines in one insertion become separate paragraphs, at the very end of the body.
    id: '11-postal-address',
    prompt: 'At the end, add our postal address on three lines: Maple Street Dental Clinic, 12 Maple Street, Springfield.',
    replay: [
      inspect('paper copy', { includeDocumentEnd: true }),
      apply([{ type: 'append_paragraph', target: { $end: true }, text: ADDRESS.join('\n') }]),
      done
    ],
    checks: [
      { type: 'adjacent', label: 'the address starts after the paper-copy line', first: { regex: 'paper copy', flags: 'i' }, second: { equals: ADDRESS[0] } },
      { type: 'adjacent', label: 'the second address line is its own paragraph', first: { equals: ADDRESS[0] }, second: { equals: ADDRESS[1] } },
      { type: 'adjacent', label: 'the third address line is its own paragraph', first: { equals: ADDRESS[1] }, second: { equals: ADDRESS[2] } },
      { type: 'text', label: 'the address lines are not joined with slashes', text: { includes: 'Clinic / 12' }, absent: true }
    ]
  },
  {
    // Several list items in one insertion join the target's list, in order.
    id: '12-two-uses',
    prompt: 'Under How We Use Your Information, after the payments bullet, add two bullets: one about meeting our legal duties and one about improving our services.',
    replay: [
      inspect('To process payments'),
      apply([{ type: 'insert_list_item', target: handle({ equals: USES[1] }), text: NEW_USES.join('\n') }]),
      done
    ],
    checks: [
      { type: 'list', label: 'both new bullets join the uses list after the payments bullet', items: [USES[0], USES[1], { regex: 'legal', flags: 'i' }, { regex: 'improv', flags: 'i' }, USES[2]], numFmt: 'bullet' }
    ]
  },
  {
    // Drafted prose of several paragraphs (like recitals) lands as separate paragraphs, not one run-on block.
    id: '13-security-section',
    prompt: 'Before Your Rights, add a short section headed How We Keep Your Information Safe, with one paragraph on how we store records and one on how long we keep them.',
    replay: [
      // Plain insertion beside a heading is refused, so the section goes after the sharing paragraph.
      inspect('We share your information only'),
      apply([{ type: 'insert_paragraph', target: handle(SHARING), position: 'after', text: SECURITY.join('\n') }]),
      done
    ],
    checks: [
      { type: 'adjacent', label: 'the section starts after the sharing paragraph', first: SHARING, second: { equals: SECURITY[0] } },
      { type: 'adjacent', label: 'the heading is its own paragraph', first: { equals: SECURITY[0] }, second: { startsWith: 'We store' } },
      { type: 'adjacent', label: 'each body paragraph is its own paragraph', first: { startsWith: 'We store' }, second: { startsWith: 'We keep' } },
      { type: 'adjacent', label: 'Your Rights follows the new section', first: { startsWith: 'We keep' }, second: { equals: 'Your Rights' } }
    ]
  }
];

export const scenario = {
  name: 'privacy-direct',
  protocol: 'direct-document-tools',
  fixture: 'tests/fixtures/golden/privacy-notice.docx',
  steps: steps.map(step => {
    const base = { id: step.id, prompt: step.prompt };
    if (step.scopeDisposition) {
      return {
        ...base,
        scopeDisposition: step.scopeDisposition,
        persist: false,
        replay: [{ chat: { text: step.refusal } }],
        checks: [{ type: 'scopeDisposition', label: 'unsupported operation explicitly refused without changing the document' }]
      };
    }
    return { ...base, replay: step.replay, checks: [...step.checks, batchEvidence] };
  })
};
