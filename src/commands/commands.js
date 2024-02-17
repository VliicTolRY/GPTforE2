/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event) {
  // Your code goes here

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
function updateSelectedModel(newModel) {
  selectedModel = newModel;
  console.log(`Model updated to: ${selectedModel}`);
}
Office.actions.associate("updateSelectedModel", updateSelectedModel);
// Register the function with Office.
Office.actions.associate("action", action);