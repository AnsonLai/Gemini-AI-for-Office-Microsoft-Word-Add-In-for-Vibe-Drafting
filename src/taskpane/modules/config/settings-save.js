// Storage writes for the Settings fields that save as soon as they change.
// `storage` is localStorage (or a stand-in with the same setItem/removeItem).

/** Save the API key. An empty key removes it, which returns the add-in to its no-key state. */
export function saveApiKeySetting(storage, value) {
  const apiKey = String(value ?? '').trim();
  if (apiKey) storage.setItem('geminiApiKey', apiKey);
  else storage.removeItem('geminiApiKey');
  return { stored: apiKey !== '' };
}

/** Save one model. `readModel(type)` throws for an invalid id, which is then not saved. */
export function saveModelSettingFrom(storage, type, readModel) {
  let model;
  try {
    model = readModel(type);
  } catch (error) {
    return { ok: false, error: error.message };
  }
  storage.setItem(type === 'slow' ? 'geminiModelSlow' : 'geminiModelFast', model);
  return { ok: true };
}

export const WEB_SEARCH_LIMIT_DEFAULT = 5;
export const WEB_SEARCH_LIMIT_MAX = 25;

/** The saved web-search limit per request, or the default when none (or an invalid one) is stored. */
export function loadWebSearchLimit(storage) {
  const stored = Number(storage.getItem('webSearchLimit'));
  return Number.isInteger(stored) && stored >= 1 && stored <= WEB_SEARCH_LIMIT_MAX ? stored : WEB_SEARCH_LIMIT_DEFAULT;
}

/** Save the web-search limit. Anything but a whole number from 1 to the maximum is refused and not saved. */
export function saveWebSearchLimitSetting(storage, value) {
  const text = String(value ?? '').trim();
  const limit = Number(text);
  if (!/^\d+$/.test(text) || limit < 1 || limit > WEB_SEARCH_LIMIT_MAX) {
    return { ok: false, error: `Enter a whole number from 1 to ${WEB_SEARCH_LIMIT_MAX}.` };
  }
  storage.setItem('webSearchLimit', String(limit));
  return { ok: true, value: limit };
}

export function saveSystemMessageSetting(storage, value) {
  storage.setItem('geminiSystemMessage', String(value ?? ''));
}
