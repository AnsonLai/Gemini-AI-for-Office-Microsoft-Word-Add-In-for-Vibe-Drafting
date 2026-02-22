import assert from 'assert';

import {
    setDefaultAuthor as setDefaultAuthorStandalone,
    getDefaultAuthor as getDefaultAuthorStandalone,
    setPlatform as setPlatformStandalone,
    getPlatform as getPlatformStandalone
} from '../src/taskpane/modules/reconciliation/standalone.js';

import {
    setDefaultAuthor as setDefaultAuthorIndex,
    getDefaultAuthor as getDefaultAuthorIndex,
    setPlatform as setPlatformIndex,
    getPlatform as getPlatformIndex
} from '../src/taskpane/modules/reconciliation/index.js';

function run() {
    setDefaultAuthorStandalone('ConfigTestAuthor');
    assert.strictEqual(getDefaultAuthorStandalone(), 'ConfigTestAuthor');
    assert.strictEqual(getDefaultAuthorIndex(), 'ConfigTestAuthor');

    setDefaultAuthorIndex('IndexAuthor');
    assert.strictEqual(getDefaultAuthorStandalone(), 'IndexAuthor');
    assert.strictEqual(getDefaultAuthorIndex(), 'IndexAuthor');

    setPlatformStandalone('OfficeOnline');
    assert.strictEqual(getPlatformStandalone(), 'OfficeOnline');
    assert.strictEqual(getPlatformIndex(), 'OfficeOnline');

    setPlatformIndex('Win32');
    assert.strictEqual(getPlatformStandalone(), 'Win32');
    assert.strictEqual(getPlatformIndex(), 'Win32');

    console.log('PASS: reconciliation config exports');
}

try {
    run();
} catch (err) {
    console.error('FAIL:', err.message);
    process.exit(1);
}
