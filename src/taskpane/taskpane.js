/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(() => {
  document.getElementById('model-selector').addEventListener('change', (event) => {
      // Save the selected model in local storage
      localStorage.setItem('SelectedModel', event.target.value);
      console.log('Model updated successfully to', event.target.value)
  });
});

function updateApiKey() {
  var apiKey = document.getElementById('openai-api-key').value;
  if(apiKey) {
      // Save the API key in local storage
      localStorage.setItem('OpenAI_ApiKey', apiKey);
      console.log('API key updated successfully.');
  } else {
      console.error('No API key entered.');
  }
}

// Event listener for the 'update-api-key' button
document.getElementById('update-api-key').addEventListener('click', updateApiKey);

