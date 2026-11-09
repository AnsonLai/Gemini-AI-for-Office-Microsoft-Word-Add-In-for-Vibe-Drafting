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

export function saveSystemMessageSetting(storage, value) {
  storage.setItem('geminiSystemMessage', String(value ?? ''));
}
