import assert from 'assert';
import {
    buildPromptParagraphSections,
    buildFormattingDiagnostics,
    isFormattingRemovalIntent,
    buildFormattingRemovalFallbackCandidate
} from '@gsd/docx-reconciliation/services/browser-demo-prompt-context.js';

function testBuildPromptParagraphSectionsIncludesFormattingOnlyWhenDifferent() {
    const sections = buildPromptParagraphSections([
        { index: 48, text: 'Section 8. British Columbia', formattedText: 'Section 8. **British Columbia**' },
        { index: 49, text: 'Standard plain paragraph', formattedText: 'Standard plain paragraph' }
    ]);

    assert.ok(sections.plainListing.includes('[P48] Section 8. British Columbia'));
    assert.ok(sections.plainListing.includes('[P49] Standard plain paragraph'));
    assert.ok(sections.formattingListing.includes('[P48_FMT] Section 8. **British Columbia**'));
    assert.ok(!sections.formattingListing.includes('[P49_FMT]'));
}

function testBuildPromptParagraphSectionsFallsBackToPlainText() {
    const sections = buildPromptParagraphSections([
        { index: 7, text: 'No markdown projection available' }
    ]);

    assert.ok(sections.plainListing.includes('[P7] No markdown projection available'));
    assert.strictEqual(sections.formattingListing, '');
}

function testBuildPromptParagraphSectionsHandlesEmptyInput() {
    const sections = buildPromptParagraphSections(null);
    assert.strictEqual(sections.plainListing, '');
    assert.strictEqual(sections.formattingListing, '');
}

function testBuildFormattingDiagnosticsFindsQuotedPhraseMatches() {
    const diagnostics = buildFormattingDiagnostics(
        [
            { index: 48, text: 'Section 8. British Columbia', formattedText: 'Section 8. **British Columbia**' },
            { index: 49, text: 'Another paragraph', formattedText: 'Another paragraph' }
        ],
        'Please unbold "British Columbia" in section 8.'
    );

    assert.ok(diagnostics.queries.includes('British Columbia'));
    assert.strictEqual(diagnostics.matches.length, 1);
    assert.strictEqual(diagnostics.matches[0].index, 48);
    assert.strictEqual(diagnostics.matches[0].differs, true);
    assert.ok(diagnostics.logLines[0].includes('P48'));
    assert.ok(diagnostics.logLines[0].includes('differs=true'));
}

function testBuildFormattingDiagnosticsHandlesNoMatches() {
    const diagnostics = buildFormattingDiagnostics(
        [
            { index: 1, text: 'Completely unrelated paragraph', formattedText: 'Completely unrelated paragraph' }
        ],
        'Please unbold "British Columbia".'
    );

    assert.ok(diagnostics.queries.includes('British Columbia'));
    assert.strictEqual(diagnostics.matches.length, 0);
    assert.ok(diagnostics.logLines[0].includes('No paragraph matches found'));
}

function testBuildFormattingDiagnosticsFindsLowercaseUnquotedPhrase() {
    const diagnostics = buildFormattingDiagnostics(
        [
            { index: 48, text: 'Section 8. British Columbia', formattedText: 'Section 8. **British Columbia**' },
            { index: 10, text: 'Governing law applies.', formattedText: 'Governing law applies.' }
        ],
        'please unbold british columbia in section 8'
    );

    assert.ok(diagnostics.queries.includes('british columbia'));
    assert.strictEqual(diagnostics.matches.length >= 1, true);
    assert.strictEqual(diagnostics.matches[0].index, 48);
}

function testIsFormattingRemovalIntentDetectsCommonCommands() {
    assert.strictEqual(isFormattingRemovalIntent('please unbold british columbia'), true);
    assert.strictEqual(isFormattingRemovalIntent('remove formatting from this line'), true);
    assert.strictEqual(isFormattingRemovalIntent('review for market standards'), false);
}

function testBuildFormattingRemovalFallbackCandidateUsesAssistantQuotedPhrase() {
    const candidate = buildFormattingRemovalFallbackCandidate(
        [
            { index: 47, text: '8. GOVERNING LAW', formattedText: '8. GOVERNING LAW' },
            { index: 48, text: 'This Agreement is governed by the laws of British Columbia.', formattedText: 'This Agreement is governed by the laws of **British Columbia**.' }
        ],
        'please unbold brittish columbia in governing law section',
        'It looks like "British Columbia" is not bolded in the Governing Law section.'
    );

    assert.ok(candidate);
    assert.strictEqual(candidate.targetRef, 48);
    assert.ok(candidate.target.includes('British Columbia'));
    assert.strictEqual(candidate.modified, candidate.target);
}

function testBuildFormattingRemovalFallbackCandidateReturnsNullWithoutFormattingDiff() {
    const candidate = buildFormattingRemovalFallbackCandidate(
        [
            { index: 12, text: 'Plain paragraph only', formattedText: 'Plain paragraph only' }
        ],
        'remove formatting from plain paragraph',
        'No formatting found.'
    );
    assert.strictEqual(candidate, null);
}

function testBuildFormattingRemovalFallbackCandidateUsesNearbySectionParagraph() {
    const candidate = buildFormattingRemovalFallbackCandidate(
        [
            { index: 47, text: '8. GOVERNING LAW', formattedText: '8. GOVERNING LAW' },
            { index: 48, text: 'This Agreement is governed by the laws of British Columbia.', formattedText: 'This Agreement is governed by the laws of **British Columbia**.' }
        ],
        'please unbold in governing law section',
        'No formatting was detected for the heading.'
    );

    assert.ok(candidate);
    assert.strictEqual(candidate.targetRef, 48);
}

function run() {
    testBuildPromptParagraphSectionsIncludesFormattingOnlyWhenDifferent();
    testBuildPromptParagraphSectionsFallsBackToPlainText();
    testBuildPromptParagraphSectionsHandlesEmptyInput();
    testBuildFormattingDiagnosticsFindsQuotedPhraseMatches();
    testBuildFormattingDiagnosticsHandlesNoMatches();
    testBuildFormattingDiagnosticsFindsLowercaseUnquotedPhrase();
    testIsFormattingRemovalIntentDetectsCommonCommands();
    testBuildFormattingRemovalFallbackCandidateUsesAssistantQuotedPhrase();
    testBuildFormattingRemovalFallbackCandidateReturnsNullWithoutFormattingDiff();
    testBuildFormattingRemovalFallbackCandidateUsesNearbySectionParagraph();
    console.log('PASS: browser demo prompt context tests');
}

try {
    run();
} catch (error) {
    console.error('FAIL:', error?.message || error);
    process.exit(1);
}

