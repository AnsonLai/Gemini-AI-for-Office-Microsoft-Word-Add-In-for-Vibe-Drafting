/** Add the refreshed document view to the same valid function-response turn. */
export function appendRefreshedDocumentContext(functionResponses, formattedText) {
  if (!Array.isArray(functionResponses)) {
    throw new TypeError('Function responses must be an array.');
  }
  if (typeof formattedText !== 'string') {
    throw new TypeError('Refreshed document context must be text.');
  }

  functionResponses.push({
    text: `Current document context after the confirmed edit:\n"""${formattedText}"""`
  });
  return functionResponses;
}
