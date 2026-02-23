import assert from 'assert';

import {
    setDefaultAuthor as setDefaultAuthorCore,
    getDefaultAuthor as getDefaultAuthorCore,
    setPlatform as setPlatformCore,
    getPlatform as getPlatformCore
} from '@ansonlai/docx-redline-js/index.js';

import {
    setDefaultAuthor as setDefaultAuthorWordEntry,
    getDefaultAuthor as getDefaultAuthorWordEntry,
    setPlatform as setPlatformWordEntry,
    getPlatform as getPlatformWordEntry
} from '../src/taskpane/modules/docx-redline-js-integration/index.js';

function run() {
    setDefaultAuthorCore('ConfigTestAuthor');
    assert.strictEqual(getDefaultAuthorCore(), 'ConfigTestAuthor');
    assert.strictEqual(getDefaultAuthorWordEntry(), 'ConfigTestAuthor');

    setDefaultAuthorWordEntry('WordEntryAuthor');
    assert.strictEqual(getDefaultAuthorCore(), 'WordEntryAuthor');
    assert.strictEqual(getDefaultAuthorWordEntry(), 'WordEntryAuthor');

    setPlatformCore('OfficeOnline');
    assert.strictEqual(getPlatformCore(), 'OfficeOnline');
    assert.strictEqual(getPlatformWordEntry(), 'OfficeOnline');

    setPlatformWordEntry('Win32');
    assert.strictEqual(getPlatformCore(), 'Win32');
    assert.strictEqual(getPlatformWordEntry(), 'Win32');

    console.log('PASS: reconciliation config exports');
}

try {
    run();
} catch (err) {
    console.error('FAIL:', err.message);
    process.exit(1);
}


