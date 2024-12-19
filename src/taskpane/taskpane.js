/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    // Check if the mailbox supports the necessary API
    if (Office.context.requirements.isSetSupported("Mailbox", "1.5")) {
      // Add a custom category
      Office.context.mailbox.masterCategories.addAsync(
        [{ displayName: "My Custom Category", color: "Preset0" }], // 'Preset0' is a predefined color
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Custom category added successfully.");
          } else {
            console.error("Failed to add custom category:", result.error.message);
          }
        }
      );
    } else {
      console.error("Required Mailbox permission set is not supported.");
    }

    // Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
    //   if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    //     const masterCategories = asyncResult.value;
    //     if (masterCategories && masterCategories.length > 0) {
    //       // Grab the first category from the master list.
    //       const categoryToAdd = [masterCategories[0].displayName];
    //       Office.context.mailbox.item.categories.addAsync(categoryToAdd, function (asyncResult) {
    //         if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    //           console.log(`Successfully assigned category '${categoryToAdd}' to item.`);
    //         } else {
    //           console.log("categories.addAsync call failed with error: " + asyncResult.error.message);
    //         }
    //       });
    //     } else {
    //       console.log("There are no categories in the master list on this mailbox. You can add categories using Office.context.mailbox.masterCategories.addAsync.");
    //     }
    //   } else {
    //     console.error(asyncResult.error);
    //   }
    // });
    // console.log("test")
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemSend, function (event) {
      console.log("ItemSend event triggered");
      showEmailMessage();
    });

    // Listen for when an item (email) is opened or selected
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function (event) {
      if (event.status === "succeeded") {
        // Display a custom HTML message when an email is selected
        showEmailMessage();
      } else {
        console.error("Failed to detect email change:", event.error);
      }
    });
  }
});

function showEmailMessage() {
  // Get the current email item
  const item = Office.context.mailbox.item;

  // Create a custom HTML message with email details
  const messageHtml = `
    <h3>New Email Selected!</h3>
    <p><strong>Subject:</strong> ${item.subject}</p>
    <p><strong>From:</strong> ${item.from.emailAddress}</p>
    <p><strong>Received:</strong> ${item.dateTimeReceived}</p>
  `;

  // Insert the HTML message into an element in the task pane
  const messageContainer = document.getElementById("email-message-container");
  if (messageContainer) {
    messageContainer.innerHTML = messageHtml;
  } else {
    console.error("Could not find the email message container.");
  }
}


export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
  showEmailMessage();
}
