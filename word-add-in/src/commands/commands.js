/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, console, fetch, setInterval, process */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
  console.log("Office.js is ready");
  setupDocumentEventHandlers();
});

/**
 * Sets up event handlers for document changes and saves
 */
function setupDocumentEventHandlers() {
  // Only proceed if we're in Word
  if (Office.context.host === Office.HostType.Word) {
    console.log("Setting up document event handlers for Word");

    // Listen for document changes
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onDocumentChanged, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Document change handler registered successfully");
      } else {
        console.error("Failed to register document change handler:", result.error);
      }
    });

    // For saves, we'll need to use a different approach since there's no direct save event
    // We'll set up a periodic check instead
    setInterval(checkForDocumentSave, 5000); // Check every 5 seconds

    console.log("Document event handlers registered");
  } else {
    console.log("Not in Word context, current host:", Office.context.host);
  }
}

// Variable to track document state for save detection
let lastDocumentState = null;

/**
 * Periodically checks if the document has been saved by comparing its state
 */
function checkForDocumentSave() {
  console.log("Checking for document save...");

  // Get document properties to check if saved
  Office.context.document.getFilePropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const docProps = result.value;
      const currentUrl = docProps.url || "";

      // If the URL has changed or we now have a URL where we didn't before,
      // we can assume a save occurred
      if (lastDocumentState === null) {
        lastDocumentState = currentUrl;
      } else if (lastDocumentState !== currentUrl && currentUrl !== "") {
        console.log("Document saved detected: URL changed from", lastDocumentState, "to", currentUrl);
        lastDocumentState = currentUrl;

        // Send save notification to server
        sendDocumentUpdate({
          timestamp: new Date().toISOString(),
          documentName: currentUrl,
          eventType: "save",
        });
      }

      // Also check for changes in the document's dirty state if that API is available
      if (Office.context.document.settings) {
        const isDirty = Office.context.document.settings.get("documentDirty") || false;
        if (!isDirty && lastDocumentState !== null) {
          // Document was dirty and is now clean, likely saved
          console.log("Document saved detected: dirty state changed");

          sendDocumentUpdate({
            timestamp: new Date().toISOString(),
            documentName: currentUrl,
            eventType: "save",
          });

          Office.context.document.settings.set("documentDirty", false);
          Office.context.document.settings.saveAsync();
        }
      }
    }
  });
}

/**
 * Handles document content changes
 * @param {Office.DocumentSelectionChangedEventArgs} eventArgs
 */
function onDocumentChanged(eventArgs) {
  console.log("Document change detected", eventArgs);

  // Get document properties
  Office.context.document.getFilePropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const docProps = result.value;
      console.log("Document properties:", docProps);

      // Mark document as dirty for save detection
      if (Office.context.document.settings) {
        Office.context.document.settings.set("documentDirty", true);
        Office.context.document.settings.saveAsync();
      }

      // Send update to server
      sendDocumentUpdate({
        timestamp: new Date().toISOString(),
        documentName: docProps.url || "Untitled",
        eventType: "change",
      });
    }
  });
}

/**
 * Sends document update data to the Node.js server
 * @param {Object} updateData - Data about the document update
 */
function sendDocumentUpdate(updateData) {
  console.log("Sending document update to server:", updateData);

  fetch(`http://${process.env.REACT_APP_HOST}:${process.env.REACT_APP_NODE_SERVER_PORT}/api/document-update`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(updateData),
  })
    .then((response) => {
      console.log("Server response status:", response.status);
      return response.json();
    })
    .then((data) => {
      console.log("Update sent successfully:", data);
    })
    .catch((error) => {
      console.error("Error sending update:", error);
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
  Office.context.mailbox.item?.notificationMessages.replaceAsync("ActionPerformanceNotification", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
