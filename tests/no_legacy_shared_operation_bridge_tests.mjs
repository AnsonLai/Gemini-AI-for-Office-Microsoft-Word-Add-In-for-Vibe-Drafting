import assert from 'assert';
import { existsSync } from 'fs';

const legacyBridgePath = 'src/taskpane/modules/commands/shared-operation-bridge.js';

function run() {
    assert.strictEqual(
        existsSync(legacyBridgePath),
        false,
        `Legacy command-layer bridge should be removed: ${legacyBridgePath}`
    );
    console.log('PASS: no legacy shared operation bridge test');
}

run();
