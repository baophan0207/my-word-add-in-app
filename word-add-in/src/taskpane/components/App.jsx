import PropTypes from "prop-types";
import * as React from "react";

import Header from "./Header";
import Button from "./BasicComponents/Button/Button";
import Popup from "./BasicComponents/Popup/Popup";
import Enhancements from "./Enhancements/Enhancements";

import "./App.scss";

/* global Word console, Office, fetch, setTimeout, FormData, Blob */

const App = (props) => {
  const { title } = props;
  const [documentStatus, setDocumentStatus] = React.useState("Monitoring document...");
  const [isPopupOpen, setIsPopupOpen] = React.useState(false);

  // Function to determine which status style to apply
  const getStatusStyle = () => {
    if (documentStatus.includes("Error")) {
      return "error-status";
    } else if (documentStatus.includes("Document updated")) {
      return "success-status";
    } else {
      return "monitoring-status";
    }
  };

  // Function to send document update to the API
  const sendDocumentUpdate = async (eventType) => {
    try {
      // Get the current document context
      await Word.run(async (context) => {
        // Get the whole document body
        const body = context.document.body;
        body.load("text");

        await context.sync();

        const contentLength = body.text.length;
        const documentName = Office.context.document.url || "Unknown Document";
        const fileName = documentName.split("/").pop() || "document.docx";

        // Only process content changes and saves
        if (eventType === "selection-change") {
          setDocumentStatus("Preparing document for upload...");

          // Create metadata object
          const metadata = {
            timestamp: new Date().toISOString(),
            documentName,
            contentLength,
            eventType,
          };

          // Get and upload the document with all its slices
          getAndUploadCompleteDocument(fileName, metadata);
        } else {
          // For selection changes, just log but don't upload
          console.log("Selection changed, not uploading document");
          setDocumentStatus("Selection changed - monitoring document");
          setTimeout(() => setDocumentStatus("Monitoring document..."), 3000);
        }
      });
    } catch (error) {
      console.error("Error sending document update:", error);
      setDocumentStatus(`⚠️ Error: ${error.message}`);
    }
  };

  // Function to get and upload the complete document
  const getAndUploadCompleteDocument = (fileName, metadata) => {
    try {
      // Get the document file
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 4194304 }, // 4MB slice size
        function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const file = result.value;
            const sliceCount = file.sliceCount;
            const slicesData = [];

            setDocumentStatus(`Retrieving document (0/${sliceCount} slices)...`);

            // Function to get a specific slice
            function getSlice(sliceIndex) {
              file.getSliceAsync(sliceIndex, function (sliceResult) {
                if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                  slicesData.push(sliceResult.value.data);
                  setDocumentStatus(`Retrieving document (${sliceIndex + 1}/${sliceCount} slices)...`);

                  // If not the last slice, get the next one
                  if (sliceIndex < sliceCount - 1) {
                    getSlice(sliceIndex + 1);
                  } else {
                    // All slices collected, now upload
                    setDocumentStatus(`Uploading document (${slicesData.length} slices)...`);
                    uploadCompleteDocument(slicesData, fileName, metadata);
                    file.closeAsync();
                  }
                } else {
                  console.error("Error getting slice:", sliceResult.error);
                  setDocumentStatus(`⚠️ Error getting slice: ${sliceResult.error.message}`);
                  file.closeAsync();
                }
              });
            }

            // Start getting slices from index 0
            getSlice(0);
          } else {
            console.error("Error getting file:", result.error);
            setDocumentStatus(`⚠️ Error getting file: ${result.error.message}`);
          }
        }
      );
    } catch (error) {
      console.error("Error extracting document:", error);
      setDocumentStatus(`⚠️ Error extracting document: ${error.message}`);
    }
  };

  // Function to upload the complete document
  const uploadCompleteDocument = (slicesData, fileName, metadata) => {
    try {
      // Calculate total length of all slices
      const totalLength = slicesData.reduce((sum, slice) => sum + slice.length, 0);
      const combined = new Uint8Array(totalLength);

      // Combine all slices into one Uint8Array
      let offset = 0;
      slicesData.forEach((slice) => {
        combined.set(new Uint8Array(slice), offset);
        offset += slice.length;
      });

      // Create a blob from the combined data with the correct MIME type
      const blob = new Blob([combined], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      // Create FormData to send the file and metadata
      const formData = new FormData();
      formData.append("document", blob, fileName);
      formData.append("metadata", JSON.stringify(metadata));

      // Send to server
      fetch("http://localhost:3001/api/upload-document-with-metadata", {
        method: "POST",
        body: formData,
      })
        .then((response) => {
          if (!response.ok) {
            throw new Error("Network response was not ok");
          }
          return response.json();
        })
        .then((data) => {
          console.log("Document uploaded successfully:", data);
          setDocumentStatus(`✓ Document updated - ${new Date().toLocaleTimeString()}`);
          setTimeout(() => setDocumentStatus("Monitoring document..."), 3000);
        })
        .catch((error) => {
          console.error("Error uploading document:", error);
          setDocumentStatus(`⚠️ Error uploading: ${error.message}`);
        });
    } catch (error) {
      console.error("Error preparing upload:", error);
      setDocumentStatus(`⚠️ Error preparing upload: ${error.message}`);
    }
  };

  // Set up document change event listeners
  React.useEffect(() => {
    const initialize = async () => {
      try {
        // Set up event handlers for document changes
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () =>
          sendDocumentUpdate("selection-change")
        );

        // Listen for content changes
        Office.context.document.addHandlerAsync("documentChange", () => sendDocumentUpdate("content-change"));

        // Listen for document save
        Office.context.document.addHandlerAsync("documentSaved", () => sendDocumentUpdate("save"));

        setDocumentStatus("Document monitoring active");
      } catch (error) {
        console.error("Error setting up document handlers:", error);
        setDocumentStatus(`⚠️ Setup error: ${error.message}`);
      }
    };

    // Initialize when the component mounts
    initialize();

    // Clean up event handlers when the component unmounts
    return () => {
      try {
        Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged);
        Office.context.document.removeHandlerAsync("documentChange");
        Office.context.document.removeHandlerAsync("documentSaved");
      } catch (error) {
        console.error("Error removing document handlers:", error);
      }
    };
  }, []);

  const closePopup = () => {
    setIsPopupOpen(false);
  };

  return (
    <div className="app-root">
      {/* Top Section - Header */}
      <div className="header-section">
        <Header logo="assets/logo.svg" title={title} message="Welcome" />
      </div>

      {/* Middle Section - Centered Content */}
      <div className="middle-section">
        <div className="content-container">
          <h2 className="content-heading">Your document is automatically saved whenever you make changes.</h2>
        </div>

        <div className="tracking-status">
          <p className="tracking-message">Tracking changes ...</p>
        </div>

        <div className="button-container">
          <Button type="primary" onClick={() => setIsPopupOpen(true)}>
            Enhance with AI
          </Button>
        </div>
      </div>

      {/* Bottom Section - Document Status */}
      <div className={`document-status ${getStatusStyle()}`}>
        <p className="status-message">{documentStatus}</p>
      </div>

      <Popup open={isPopupOpen} onClose={closePopup} logo="assets/logo.svg" name="IP Agent AI">
        <Enhancements />
      </Popup>
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
