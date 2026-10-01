const LOCAL_COLLECTOR_HOSTS = new Set(['localhost', '127.0.0.1', '[::1]']);

/** Post diagnostic startup timings to the local development collector only. */
export async function postTaskpaneStartupProfile(record, {
  location = globalThis.window?.location,
  fetchImpl = globalThis.fetch
} = {}) {
  if (!location || !LOCAL_COLLECTOR_HOSTS.has(location.hostname) || typeof fetchImpl !== 'function') return false;

  try {
    await fetchImpl(new URL('/startup-result', location.origin), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      cache: 'no-store',
      body: JSON.stringify(record)
    });
    return true;
  } catch (error) {
    console.warn('[TaskpaneStartupProfile] Local collector post failed.', error);
    return false;
  }
}
