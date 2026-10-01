import assert from 'node:assert/strict';
import { createLazyModuleLoader } from '../../src/taskpane/modules/utils/lazy-module-loader.js';

async function testConcurrentFirstLoadAndInitialization() {
    let resolveImport;
    let loads = 0;
    let initializations = 0;
    const module = { execute: () => 'ready' };
    const loader = createLazyModuleLoader(
        () => {
            loads += 1;
            return new Promise(resolve => { resolveImport = resolve; });
        },
        loaded => {
            assert.equal(loaded, module);
            initializations += 1;
        }
    );

    const first = loader();
    const second = loader();
    await Promise.resolve();
    assert.equal(loads, 1, 'concurrent calls share one import');
    resolveImport(module);
    const [firstResult, secondResult] = await Promise.all([first, second]);
    assert.equal(firstResult, module);
    assert.equal(secondResult, module);
    assert.equal(await loader(), module, 'later calls receive the cached module');
    assert.equal(initializations, 1, 'module dependencies are initialized exactly once');
}

async function testFailedLoadCanRetry() {
    let loads = 0;
    let initializations = 0;
    const module = { execute: () => 'recovered' };
    const loader = createLazyModuleLoader(
        async () => {
            loads += 1;
            if (loads === 1) throw new Error('temporary chunk failure');
            return module;
        },
        () => { initializations += 1; }
    );

    await assert.rejects(loader(), /temporary chunk failure/);
    assert.equal(await loader(), module, 'a failed first import can be retried');
    assert.equal(loads, 2);
    assert.equal(initializations, 1);
}

try {
    await testConcurrentFirstLoadAndInitialization();
    await testFailedLoadCanRetry();
    console.log('PASS: lazy taskpane module loader tests');
} catch (error) {
    console.error('FAIL:', error?.stack || error?.message || error);
    process.exitCode = 1;
}
