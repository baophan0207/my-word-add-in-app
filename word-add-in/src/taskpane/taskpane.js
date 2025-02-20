/* global Word console, Office, document */

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
  // Enable auto-open
  const enableAutoOpen = async () => {
    try {
      // Set the setting to auto-open
      await Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
      // Save the settings
      await Office.context.document.settings.saveAsync();
      console.log("Auto-open enabled successfully");
    } catch (error) {
      console.error("Error enabling auto-open:", error);
    }
  };

  // Disable auto-open
  const disableAutoOpen = async () => {
    try {
      // Remove the auto-open setting
      await Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", false);
      // Save the settings
      await Office.context.document.settings.saveAsync();
      console.log("Auto-open disabled successfully");
    } catch (error) {
      console.error("Error disabling auto-open:", error);
    }
  };

  // Check current auto-open status
  // const checkAutoOpenStatus = () => {
  //   const isAutoOpen = Office.context.document.settings.get("Office.AutoShowTaskpaneWithDocument");
  //   console.log("Auto-open status:", isAutoOpen);
  //   return isAutoOpen;
  // };

  // Example: Add buttons to control auto-open
  document.getElementById("enableAutoOpen").onclick = enableAutoOpen;
  document.getElementById("disableAutoOpen").onclick = disableAutoOpen;
});
