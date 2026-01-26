
import assert from 'assert';

/**
 * Re-implementation of parseMarkdownList from taskpane.js for testing purposes.
 * We want to verify this logic correctly identifies "Text" vs "List Item".
 */
function parseMarkdownList(content) {
    if (!content) return null;

    const lines = content.trim().split('\n');
    const items = [];

    for (const line of lines) {
        if (!line.trim()) continue;

        // Unified marker regex matching NumberingService and Pipeline
        const markerRegex = /^(\s*)((?:\d+(?:\.\d+)*\.?|\((?:\d+|[a-zA-Z]|[ivxlcIVXLC]+)\)|[a-zA-Z]\.|\d+\.|[ivxlcIVXLC]+\.|[-*•])(?=\s|$)\s*)(.*)$/;
        const match = line.match(markerRegex);

        if (match) {
            const indent = match[1];
            const marker = match[2].trim();
            const text = match[3];
            const level = Math.floor(indent.length / 2); // 2 spaces per level

            // Determine type based on marker
            const isBullet = /^[-*•]$/.test(marker);
            items.push({
                type: isBullet ? 'bullet' : 'numbered',
                level,
                text: text.trim(),
                marker: marker
            });
            continue;
        }

        // If line doesn't match list pattern, still include as text
        items.push({ type: 'text', level: 0, text: line.trim() });
    }

    if (items.length === 0) return null;

    // Determine primary type (numbered or bullet)
    const hasNumbered = items.some(i => i.type === 'numbered');
    const hasBullet = items.some(i => i.type === 'bullet');

    return {
        type: hasNumbered ? 'numbered' : (hasBullet ? 'bullet' : 'text'),
        items: items
    };
}

// --- Tests ---

console.log('Running Mixed Content Parsing Tests...');

// Case 1: Mixed Content (Preamble + List) - The User's Bug Case
const input1 = `If the Receiving Party is required by law...
1. provide the Disclosing Party...
2. reasonably cooperate...
3. if disclosure is ultimately required...`;

const result1 = parseMarkdownList(input1);
assert.strictEqual(result1.type, 'numbered', 'Should detect presence of numbered list');
assert.strictEqual(result1.items.length, 4, 'Should have 4 items');
assert.strictEqual(result1.items[0].type, 'text', 'Item 0 should be text (preamble)');
assert.strictEqual(result1.items[1].type, 'numbered', 'Item 1 should be numbered');
assert.strictEqual(result1.items[1].marker, '1.', 'Item 1 marker check');
assert.strictEqual(result1.items[2].type, 'numbered', 'Item 2 should be numbered');
assert.strictEqual(result1.items[3].type, 'numbered', 'Item 3 should be numbered');

console.log('✅ Case 1 Passed: Mixed Content correctly parsed');

// Case 2: Pure Numbered List
const input2 = `1. Item One
2. Item Two`;

const result2 = parseMarkdownList(input2);
assert.strictEqual(result2.items[0].type, 'numbered', 'Pure list Item 0 should be numbered');
assert.strictEqual(result2.items[1].type, 'numbered', 'Pure list Item 1 should be numbered');

console.log('✅ Case 2 Passed: Pure Numbered List correctly parsed');

// Case 3: Text Only
const input3 = `Just some text.
More text.`;

const result3 = parseMarkdownList(input3);
assert.strictEqual(result3.type, 'text', 'Text only should be type text');
assert.strictEqual(result3.items[0].type, 'text');
assert.strictEqual(result3.items[1].type, 'text');

console.log('✅ Case 3 Passed: Text Only correctly parsed');

// Case 4: Bullet List
const input4 = `- Bullet 1
* Bullet 2`;

const result4 = parseMarkdownList(input4);
assert.strictEqual(result4.type, 'bullet', 'Should be bullet type');
assert.strictEqual(result4.items[0].type, 'bullet');
assert.strictEqual(result4.items[1].type, 'bullet');

console.log('✅ Case 4 Passed: Bullet List correctly parsed');

console.log('🎉 All Parsing Tests Passed!');
