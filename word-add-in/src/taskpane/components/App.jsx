import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import { tokens, makeStyles } from "@fluentui/react-components";

/* global Word console, Office, fetch, setTimeout */

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

        const updateData = {
          timestamp: new Date().toISOString(),
          documentName,
          contentLength,
          eventType,
        };

        // API endpoint from your server
        const apiUrl = `http://localhost:3001/api/document-update`;

        // Send the update
        const response = await fetch(apiUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify(updateData),
        });

        if (response.ok) {
          setDocumentStatus(`✓ Document updated - ${new Date().toLocaleTimeString()}`);
          // Reset status message after 3 seconds
          setTimeout(() => setDocumentStatus("Monitoring document..."), 3000);
        } else {
          throw new Error("Failed to send update");
        }
      });
    } catch (error) {
      console.error("Error sending document update:", error);
      setDocumentStatus(`⚠️ Error: ${error.message}`);
    }
  };

  // Set up document change event listeners
  React.useEffect(() => {
    console.log("useEffect");
    const initialize = async () => {
      try {
        // Set up event handlers for document changes
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () =>
          sendDocumentUpdate("selection-change")
        );
        console.log("selection-change");

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
