import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import HeroList from "./HeroList";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

/* global Office, Word, console, window*/

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  const [previousContent, setPreviousContent] = React.useState("");
  const [isAutoShowEnabled, setIsAutoShowEnabled] = React.useState(false);
  const [updateStatus, setUpdateStatus] = React.useState("");

  React.useEffect(() => {
    Office.onReady(async () => {
      try {
        // Check current auto-show setting
        const settings = Office.context.document.settings;
        const currentSetting = settings.get("Office.AutoShowTaskpaneWithDocument");
        setIsAutoShowEnabled(!!currentSetting);

        const isInstalled = settings.get("AddInInstallationStatus");

        if (!isInstalled) {
          // Hiển thị dialog yêu cầu cài đặt
          // Office.context.ui.displayDialogAsync(
          //   "https://your-domain.com/install-prompt.html",
          //   { height: 30, width: 20 },
          //   function (result) {
          //     if (result.status === Office.AsyncResultStatus.Succeeded) {
          //       const dialog = result.value;
          //       dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
          //     }
          //   }
          // );
          console.log("Add-in not installed");
        } else {
          // Add-in đã được cài đặt, tự động kích hoạt
          await Office.addin.showAsTaskpane();

          // Lưu trạng thái auto-show
          settings.set("Office.AutoShowTaskpaneWithDocument", true);
          await settings.saveAsync();
        }

        // Enable auto-show for this document
        settings.set("Office.AutoShowTaskpaneWithDocument", true);
        await settings.saveAsync();
        setIsAutoShowEnabled(true);
        console.log("Auto-show enabled for this document");

        // Set up document change event handler
        Office.context.document.addHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          () => checkContentChanges(),
          (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.error("Failed to add document change handler:", result.error);
            }
          }
        );
      } catch (error) {
        console.error("Error setting up document:", error);
      }
    });

    // Cleanup handler on unmount
    return () => {
      if (Office.context.document) {
        Office.context.document.removeHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          { handler: checkContentChanges },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.error("Failed to remove document change handler:", result.error);
            }
          }
        );
      }
    };
  }, []);

  const sendUpdateToServer = async (previousLength, currentLength) => {
    try {
      const response = await window.fetch("http://localhost:3001/api/document-update", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          timestamp: new Date().toISOString(),
          previousLength,
          currentLength,
        }),
      });

      const data = await response.json();
      console.log("data", data);
      setUpdateStatus("Update sent to server successfully");
      window.setTimeout(() => setUpdateStatus(""), 3000); // Clear status after 3 seconds
    } catch (error) {
      console.error("Error sending update to server:", error);
      setUpdateStatus("Failed to send update to server");
    }
  };

  const checkContentChanges = async () => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();

        const currentContent = body.text;

        if (currentContent !== previousContent) {
          console.log("Content changed at:", new Date().toLocaleTimeString());
          console.log("Previous content length:", previousContent.length);
          console.log("Current content length:", currentContent.length);
          console.log("Content changed!");

          // Send update to server
          await sendUpdateToServer(previousContent.length, currentContent.length);

          setPreviousContent(currentContent);
        }
      });
    } catch (error) {
      console.error("Error checking content:", error);
    }
  };

  // Function to toggle auto-show
  const toggleAutoShow = async () => {
    try {
      const settings = Office.context.document.settings;
      const newValue = !isAutoShowEnabled;

      settings.set("Office.AutoShowTaskpaneWithDocument", newValue);
      await settings.saveAsync();
      setIsAutoShowEnabled(newValue);
      console.log(`Auto-show ${newValue ? "enabled" : "disabled"}`);
    } catch (error) {
      console.error("Error toggling auto-show:", error);
    }
  };

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

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={title} message="Welcome" />
      <div style={{ padding: "10px", backgroundColor: "white" }}>
        <div>Document Status: {previousContent ? "Content has changed" : "No changes"}</div>
        {updateStatus && <div style={{ color: "green", marginTop: "5px" }}>{updateStatus}</div>}
        <div style={{ marginTop: "10px" }}>
          <label>
            <input type="checkbox" checked={isAutoShowEnabled} onChange={toggleAutoShow} /> Auto-open with document
          </label>
        </div>
      </div>
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} />
      <TextInsertion insertText={insertText} />
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
