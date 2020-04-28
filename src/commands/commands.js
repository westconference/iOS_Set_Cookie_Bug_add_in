/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

const DIALOG_INVALID_PAGE = 12002;
const DIALOG_CLOSED_ERROR_CODE = 12006;

const DIALOG_URL = 'https://localhost:3000/dialog.html';

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  openDialog(DIALOG_URL)
    .then((message) => {
      console.log('Received message: ' + message);
      console.log('Dialog completed');
    }).catch((error) => {
      console.log('Failure during dialog: ' + error);
    }).finally(() => {
      event.completed();
    });
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

function openDialog (openUrl) {
  return new Promise((resolve, reject) => {
    debugger;
    Office.context.ui.displayDialogAsync(openUrl, { width: 40, height: 30, promptBeforeOpen: false }, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        reject('Failure opening dialog');
        return;
      }
      var dialog = asyncResult.value;

      // Handle auth data sent from Dialog
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (data) {
        console.log('Closing auth dialog');
        dialog.close();
        resolve(data.message);
      });

      // Handle User actions/errors https://docs.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins#errors-and-events-in-the-dialog-window
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function (event) {
        switch (event.error) {
          case DIALOG_INVALID_PAGE:
            reject('Invalid page');
            break; // unable to find or load page, continue to clode dialog
          case DIALOG_CLOSED_ERROR_CODE:
            reject('Closed by user');
            return; // break; - dialog closed by user, no need to continue
          default:
            reject('Other error occurred with dialog.');
            break;
        }

        try {
          dialog.close();
        } catch (ex) {
          console.log('Failed to close auth dialog: ' + ex);
        }
      });
    });
  });
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;
