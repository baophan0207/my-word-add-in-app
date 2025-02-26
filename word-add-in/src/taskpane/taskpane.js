/* global Word console, Office, fetch */

export async function insertText(text) {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

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
    // Listen for document changes
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onDocumentChanged);

    console.log("Document event handlers registered");
  }
}

/**
 * Handles document content changes
 * @param {Office.DocumentSelectionChangedEventArgs} eventArgs
 */
function onDocumentChanged() {
  // Get document properties
  Office.context.document.getFilePropertiesAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const docProps = result.value;

      // Send update to server
      sendDocumentUpdate({
        timestamp: new Date().toISOString(),
        documentName: docProps.url || "Untitled",
        previousLength: 0, // We don't have previous length in this example
        currentLength: 0, // You could get content length if needed
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
  fetch(`http://localhost:3001/api/document-update`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(updateData),
  })
    .then((response) => response.json())
    .then((data) => {
      console.log("Update sent successfully:", data);
    })
    .catch((error) => {
      console.error("Error sending update:", error);
    });
}
