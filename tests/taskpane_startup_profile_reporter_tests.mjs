import assert from 'node:assert/strict';
import { postTaskpaneStartupProfile } from '../src/taskpane/modules/diagnostics/startup-profile-reporter.js';

async function testPostsOnlyToLocalCollector() {
  const calls = [];
  const record = { status: 'passed', providerCalls: false };
  const fetchImpl = async (url, options) => {
    calls.push({ url: String(url), options });
    return { ok: true };
  };

  assert.equal(await postTaskpaneStartupProfile(record, {
    location: { hostname: 'localhost', origin: 'https://localhost:3000' },
    fetchImpl
  }), true);
  assert.deepEqual(calls.map(call => call.url), ['https://localhost:3000/startup-result']);
  assert.equal(calls[0].options.method, 'POST');
  assert.equal(calls[0].options.cache, 'no-store');
  assert.deepEqual(JSON.parse(calls[0].options.body), record);

  for (const hostname of ['127.0.0.1', '[::1]']) {
    assert.equal(await postTaskpaneStartupProfile(record, {
      location: { hostname, origin: `https://${hostname}:3000` },
      fetchImpl
    }), true);
  }
  assert.equal(calls.length, 3);
}

async function testDoesNotContactNonlocalOrigins() {
  let fetches = 0;
  const result = await postTaskpaneStartupProfile({ status: 'passed' }, {
    location: { hostname: 'addin.example.com', origin: 'https://addin.example.com' },
    fetchImpl: async () => { fetches++; }
  });
  assert.equal(result, false);
  assert.equal(fetches, 0);
}

try {
  await testPostsOnlyToLocalCollector();
  await testDoesNotContactNonlocalOrigins();
  console.log('PASS: local-only taskpane startup profile reporter');
} catch (error) {
  console.error('FAIL:', error?.stack || error?.message || error);
  process.exitCode = 1;
}
