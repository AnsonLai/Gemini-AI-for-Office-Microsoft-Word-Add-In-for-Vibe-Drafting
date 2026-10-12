/* global document */

export function populateModelSetting(type, model) {
  const select = document.getElementById(`model-select-${type}`);
  const custom = document.getElementById(`model-custom-${type}`);
  const preset = Array.from(select.options).some(option => option.value === model);
  select.value = preset ? model : 'custom';
  custom.value = preset ? '' : model;
  custom.hidden = select.value !== 'custom';
  select.onchange = () => { custom.hidden = select.value !== 'custom'; };
}

export function readModelSetting(type) {
  const selected = document.getElementById(`model-select-${type}`).value;
  const model = selected === 'custom' ? document.getElementById(`model-custom-${type}`).value.trim() : selected;
  if (!/^[a-zA-Z0-9][a-zA-Z0-9._-]*$/.test(model)) {
    throw new Error('Enter a model ID (letters, numbers, dots, underscores or hyphens) for each custom model.');
  }
  return model;
}
