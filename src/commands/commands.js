/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady((info) => {
  // If needed, Office.js is ready to be called.
  
});

function addOriginalFlag(event) {
  Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const masterCategories = asyncResult.value;
      if (masterCategories && masterCategories.length > 0) {
        // Grab the first category from the master list.
        // const categoryToAdd = [masterCategories[0].displayName];
        const categoryToAdd = ["Original"];

        Office.context.mailbox.item.categories.addAsync(categoryToAdd, function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(`Successfully assigned category '${categoryToAdd}' to item.`);
          } else {
            console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
          }
        });
      } else {
        console.log("There are no categories in the master list on this mailbox. You can add categories using Office.context.mailbox.masterCategories.addAsync.");
      }
    } else {
      console.error(asyncResult.error);
    }
  });

}

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
// export function onMessageRead(event) {
//   const item = Office.context.mailbox.item;
//   console.log(`Email opened: ${item.subject}`);
//   event.completed();
// }

// Register the function with Office.
Office.actions.associate("action", action);
