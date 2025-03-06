import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import HeroList from "./HeroList";
// import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
// import { insertText } from "../taskpane";

/* global Word console, Office, fetch, setTimeout, process */

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [documentStatus, setDocumentStatus] = React.useState("Monitoring document...");
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];

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
        const apiUrl = `http://${process.env.REACT_APP_HOST || "localhost"}:${process.env.REACT_APP_NODE_SERVER_PORT || "3001"}/api/document-update`;

        // Send the update
        const response = await fetch(apiUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify(updateData),
        });

        if (response.ok) {
          setDocumentStatus(`Document ${eventType} event sent - ${new Date().toLocaleTimeString()}`);
          // Reset status message after 3 seconds
          setTimeout(() => setDocumentStatus("Monitoring document..."), 3000);
        } else {
          throw new Error("Failed to send update");
        }
      });
    } catch (error) {
      console.error("Error sending document update:", error);
      setDocumentStatus(`Error: ${error.message}`);
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
        setDocumentStatus(`Setup error: ${error.message}`);
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
      <Header logo="assets/logo-filled.png" title={title} message="Welcome" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
      {/* <TextInsertion insertText={insertText} /> */}

      <div
        style={{
          margin: "20px 0",
          padding: "10px",
          backgroundColor: "#f5f5f5",
          borderRadius: "4px",
          textAlign: "center",
        }}
      >
        <p>{documentStatus}</p>
      </div>
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
