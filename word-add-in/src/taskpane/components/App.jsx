import { makeStyles, tokens } from "@fluentui/react-components";
import PropTypes from "prop-types";
import * as React from "react";
import Header from "./Header";

/* global Word console, Office, fetch, setTimeout, FormData, Blob */

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    backgroundColor: "#212E33",
    color: "#D2D2D2",
  },
  statusContainer: {
    padding: "12px",
    textAlign: "center",
    marginTop: "10px",
    transition: "background-color 0.3s ease",
  },
  successStatus: {
    backgroundColor: "#ecf8f0",
    color: "#0e6027",
    fontWeight: tokens.fontWeightSemibold,
  },
  errorStatus: {
    backgroundColor: "#fef0f1",
    color: "#d13438",
    fontWeight: tokens.fontWeightSemibold,
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [documentStatus, setDocumentStatus] = React.useState("Monitoring document...");

  // Function to determine which status style to apply
  const getStatusStyle = () => {
    if (documentStatus.includes("Error")) {
      return styles.errorStatus;
    } else if (documentStatus.includes("Document updated")) {
      return styles.successStatus;
    } else {
      return styles.monitoringStatus;
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

  return (
    <div className={styles.root}>
      <Header logo="assets/logo.svg" title={title} message="Welcome" />

      <div
        style={{
          width: "100%",
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          justifyContent: "center",
        }}
      >
        <h2
          style={{
            fontSize: tokens.fontSizeBase500,
            fontColor: tokens.colorNeutralBackgroundStatic,
            fontWeight: tokens.fontWeightRegular,
            paddingLeft: "10px",
            paddingRight: "10px",
            lineHeight: "normal",
            textAlign: "center",
          }}
        >
          Your document will be saved automatically with better text!
        </h2>
      </div>

      <div className={`${styles.statusContainer} ${getStatusStyle()}`}>
        <p style={{ margin: 0, padding: "4px 0" }}>{documentStatus}</p>
      </div>
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
