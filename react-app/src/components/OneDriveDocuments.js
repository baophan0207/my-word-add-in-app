import React from "react";
import "./OneDriveDocument.css";
import { io } from "socket.io-client";

class OneDriveDocuments extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      documents: [],
      wordInstalled: false,
      addinInstalled: false,
      status: "",
      activeDocuments: {}, // Stores status for each active document
      socket: null,
    };
  }

  componentDidMount() {
    this.loadDocuments();
    this.connectSocket();
  }

  componentWillUnmount() {
    // Clean up socket connection
    if (this.state.socket) {
      this.state.socket.disconnect();
    }
  }

  connectSocket = () => {
    const socket = io(
      `http://${process.env.REACT_APP_HOST}:${process.env.REACT_APP_NODE_SERVER_PORT}`
    );

    socket.on("connect", () => {
      console.log("Connected to server socket");
    });

    socket.on("document-update", (update) => {
      console.log("Received document update:", update);

      // Update the active documents state with latest status
      this.setState((prevState) => ({
        activeDocuments: {
          ...prevState.activeDocuments,
          [update.documentName]: {
            lastUpdate: update.timestamp,
            status: update.eventType,
            contentLength: update.contentLength,
          },
        },
      }));

      // Also show a temporary status message
      this.setState({
        status: `Document "${this.getFileName(update.documentName)}" ${
          update.eventType === "content-change"
            ? "content changed"
            : update.eventType
        } at ${new Date(update.timestamp).toLocaleTimeString()}`,
      });

      // Clear status message after 3 seconds
      setTimeout(() => {
        this.setState({ status: "" });
      }, 3000);
    });

    socket.on("disconnect", () => {
      console.log("Disconnected from server");
    });

    this.setState({ socket });
  };

  getFileName(path) {
    // Extract just the filename from a path
    return path.split("\\").pop().split("/").pop();
  }

  loadDocuments = async () => {
    try {
      const response = await fetch(
        `http://${process.env.REACT_APP_HOST}:${process.env.REACT_APP_NODE_SERVER_PORT}/api/documents`
      );
      const documents = await response.json();
      this.setState({ documents });
    } catch (error) {
      console.error("Error loading documents:", error);
      this.setState({ status: "Error loading documents" });
    }
  };

  checkWordAndAddin = async (documentUrl) => {
    try {
      // Check Word installation
      const wordResponse = await fetch(
        `http://${process.env.REACT_APP_HOST}:${process.env.REACT_APP_NODE_SERVER_PORT}/api/check-word`
      );
      const wordData = await wordResponse.json();
      this.setState({ wordInstalled: wordData.isWordInstalled });
      console.log(this.state.wordInstalled);

      if (!wordData.isWordInstalled) {
        this.setState({
          status: "Microsoft Word is not installed. Please install Word first.",
        });
        return false;
      }

      // First open the document and then check/setup add-in
      const openResponse = await fetch(
        `http://${process.env.REACT_APP_HOST}:${process.env.REACT_APP_NODE_SERVER_PORT}/api/setup-office-addin`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ documentUrl }),
        }
      );

      if (!openResponse.ok) {
        this.setState({ status: "Failed to open document" });
        return false;
      }
      return true;
    } catch (error) {
      console.error("Error checking Word and Add-in:", error);
      this.setState({ status: "Error checking Word and Add-in installation" });
      return false;
    }
  };

  openDocument = async (doc) => {
    try {
      const baseUrl = doc.url;

      // First open the document using protocol handler
      const protocolUrl = `ms-word:ofe|u|${baseUrl}`;
      window.open(protocolUrl);

      this.setState({ status: "Opening document..." });

      // Add this document to active documents with initial state
      this.setState((prevState) => ({
        activeDocuments: {
          ...prevState.activeDocuments,
          [baseUrl]: {
            lastUpdate: new Date().toISOString(),
            status: "opening",
            contentLength: 0,
          },
        },
      }));

      // Now check Word and setup add-in with the specific document URL
      await this.checkWordAndAddin(baseUrl);

      this.setState({
        status: "Document opened successfully with add-in enabled",
      });
    } catch (error) {
      console.error("Error opening document:", error);
      this.setState({ status: "Error opening document" });
    }
  };

  getDocumentStatusClass(status) {
    switch (status) {
      case "content-change":
        return "status-changed";
      case "save":
        return "status-saved";
      case "opening":
        return "status-opening";
      default:
        return "";
    }
  }

  render() {
    return (
      <div className="documents-list">
        <h2>Your Word Documents</h2>
        {this.state.status && (
          <div
            className="status-message"
            style={{
              color: this.state.status.includes("success") ? "green" : "red",
              marginBottom: "20px",
              padding: "10px",
              backgroundColor: "#f8f8f8",
              borderRadius: "4px",
            }}
          >
            {this.state.status}
          </div>
        )}
        <div className="document-container">
          <div className="documents-grid">
            {this.state.documents.map((doc) => {
              const docStatus = this.state.activeDocuments[doc.url];
              const statusClass = docStatus
                ? this.getDocumentStatusClass(docStatus.status)
                : "";

              return (
                <div
                  key={doc.id}
                  className={`document-item ${statusClass}`}
                  onClick={() => this.openDocument(doc)}
                >
                  <span>{doc.name}</span>
                  {docStatus && (
                    <div className="document-status">
                      <small>
                        Status: {docStatus.status}
                        {docStatus.status === "content-change" &&
                          ` (${docStatus.contentLength} chars)`}
                        <br />
                        Last update:{" "}
                        {new Date(docStatus.lastUpdate).toLocaleTimeString()}
                      </small>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      </div>
    );
  }
}

export default OneDriveDocuments;
