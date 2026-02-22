import assert from 'assert';
import { resolveInsertListItemLevel } from '../src/taskpane/modules/commands/list-level-utils.js';

function run() {
    assert.deepStrictEqual(
        resolveInsertListItemLevel(1, 0),
        { baseIlvl: 1, normalizedIndent: 0, appliedIndent: 0, newIlvl: 1 },
        'same-level insert should keep ilvl'
    );

    assert.deepStrictEqual(
        resolveInsertListItemLevel(1, 1),
        { baseIlvl: 1, normalizedIndent: 1, appliedIndent: 1, newIlvl: 2 },
        'one-level deeper insert should increment ilvl by one'
    );

    assert.deepStrictEqual(
        resolveInsertListItemLevel(1, 2),
        { baseIlvl: 1, normalizedIndent: 2, appliedIndent: 1, newIlvl: 2 },
        'out-of-contract deep indent values should clamp to one level deeper'
    );

    assert.deepStrictEqual(
        resolveInsertListItemLevel(2, -3),
        { baseIlvl: 2, normalizedIndent: -3, appliedIndent: -1, newIlvl: 1 },
        'out-of-contract negative indent values should clamp to one level shallower'
    );

    assert.deepStrictEqual(
        resolveInsertListItemLevel(8, 1),
        { baseIlvl: 8, normalizedIndent: 1, appliedIndent: 1, newIlvl: 8 },
        'resolved level should remain in Word ilvl range 0..8'
    );
}

try {
    run();
    console.log('PASS: insert list item level resolution tests');
} catch (error) {
    console.error('FAIL:', error?.message || error);
    process.exit(1);
}
