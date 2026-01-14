import React from "react";
import "./WordAddIn.css";
import { io } from "socket.io-client";

class WordAddIn extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      documents: [],
      wordInstalled: false,
      addinInstalled: false,
      addinHandlerInstalled: false,
      status: "",
      activeDocuments: {}, // Stores status for each active document
      socket: null,
      pendingDocumentUri: null,
      checkingHandler: false,
      currentDocument: null,
    };
  }

  componentDidMount() {
    this.loadDocuments();
    this.connectSocket();
    this.checkAddinHandlerInstalled();
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

      return true;
    } catch (error) {
      console.error("Error checking Word and Add-in:", error);
      this.setState({ status: "Error checking Word and Add-in installation" });
      return false;
    }
  };

  checkAddinHandlerInstalled = async () => {
    this.setState({ checkingHandler: true });

    const LOCAL_PING_URL = "http://127.0.0.1:9876/ping";
    const MAX_RETRIES = 10;
    const RETRY_DELAY = 500; // 500ms between retries

    // Helper function to check if local server is responding
    const pingLocalServer = async () => {
      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 1000);

        const response = await fetch(LOCAL_PING_URL, {
          method: "GET",
          signal: controller.signal,
        });
        clearTimeout(timeoutId);

        if (response.ok) {
          const data = await response.json();
          return data.success === true;
        }
        return false;
      } catch (error) {
        // Server not responding yet
        return false;
      }
    };

    // First, check if handler is already running (from a previous check)
    const alreadyRunning = await pingLocalServer();
    if (alreadyRunning) {
      console.log("Handler already running - detected via local ping");
      this.setState({
        addinHandlerInstalled: true,
        checkingHandler: false,
      });
      return;
    }

    // Trigger the custom protocol to start the handler
    console.log("Triggering wordaddin://ping to start handler...");

    // Use a hidden link click to trigger protocol (more reliable than location.href)
    const link = document.createElement("a");
    link.href = "wordaddin://ping";
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    // Poll for local server to come up
    let detected = false;
    for (let i = 0; i < MAX_RETRIES; i++) {
      await new Promise((resolve) => setTimeout(resolve, RETRY_DELAY));

      const isRunning = await pingLocalServer();
      if (isRunning) {
        detected = true;
        console.log(`Handler detected after ${i + 1} attempts`);
        break;
      }
    }

    this.setState({
      addinHandlerInstalled: detected,
      checkingHandler: false,
    });

    if (detected) {
      console.log("Word Add-in Handler is installed and responding");
    } else {
      console.log("Word Add-in Handler not detected - may not be installed");
    }
  };

  openDocument = async (doc) => {
    // Check if handler is installed first
    if (!this.state.addinHandlerInstalled) {
      this.setState({
        status:
          "Word Add-in Handler is not installed. Please download and install it first.",
      });
      return;
    }

    try {
      const baseUrl = doc.url;

      console.log("Opening document:", baseUrl);

      // Set status to indicate document is being opened
      this.setState({
        status: `Opening document "${doc.name}" with Word Add-in...`,
      });

      // Use custom protocol to open document and setup add-in
      const customProtocolUrl = `wordaddin://open?documentUrl=${encodeURIComponent(
        baseUrl
      )}&documentName=${encodeURIComponent(doc.name)}`;

      console.log("Launching custom protocol:", customProtocolUrl);

      // Launch the custom protocol handler
      window.location.href = customProtocolUrl;

      // Add the document to active documents
      this.setState((prevState) => ({
        activeDocuments: {
          ...prevState.activeDocuments,
          [baseUrl]: {
            lastUpdate: new Date().toISOString(),
            status: "opening",
            contentLength: 0,
          },
        },
        status: `Document "${doc.name}" is being opened with Word Add-in Handler.`,
      }));

      // Clear status after a reasonable time
      setTimeout(() => {
        this.setState((prevState) => {
          if (prevState.status.includes("is being opened")) {
            return { status: "" };
          }
          return null;
        });
      }, 7000);
    } catch (error) {
      console.error("Error opening document:", error);
      this.setState({
        status: `Error opening document: ${error.message}`,
      });
    }
  };

  // Method to handle the add-in launch with direct user gesture
  handleLaunchAddin = () => {
    const { pendingDocumentUri, currentDocument } = this.state;

    if (pendingDocumentUri) {
      // This will work because it's directly triggered by a user click
      window.open(pendingDocumentUri, "_blank");

      this.setState({
        status: `Add-in setup initiated for "${currentDocument?.name}". Word should be configured shortly.`,
        pendingDocumentUri: null,
        currentDocument: null,
      });

      // Clear status after a reasonable time
      setTimeout(() => {
        this.setState((prevState) => {
          if (prevState.status.includes("Add-in setup initiated")) {
            return { status: "" };
          }
          return null;
        });
      }, 7000);
    }
  };

  downloadHandlerInstaller = () => {
    const installerUrl = `http://${process.env.REACT_APP_HOST}:${process.env.REACT_APP_NODE_SERVER_PORT}/downloads/WordAddinHandlerSetupProduction.exe`;

    const a = document.createElement("a");
    a.href = installerUrl;
    a.download = "WordAddinHandlerSetup.exe";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    this.setState({
      status:
        "Downloading Word Add-in Handler installer. Please run the installer when download completes.",
    });
  };

  handleUriLaunch = () => {
    const { pendingDocumentUri } = this.state;

    if (pendingDocumentUri) {
      window.location.href = pendingDocumentUri;

      this.setState({
        pendingDocumentUri: null,
        status: "Document opening process initiated. Word should open shortly.",
      });

      const docUrl = new URLSearchParams(pendingDocumentUri.split("?")[1]).get(
        "documentUrl"
      );
      if (docUrl) {
        this.setState((prevState) => ({
          activeDocuments: {
            ...prevState.activeDocuments,
            [docUrl]: {
              lastUpdate: new Date().toISOString(),
              status: "opening",
              contentLength: 0,
            },
          },
        }));
      }
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

        {/* Handler Installation Check */}
        {!this.state.addinHandlerInstalled && !this.state.checkingHandler && (
          <div
            className="handler-warning"
            style={{
              backgroundColor: "#fff3cd",
              color: "#856404",
              padding: "15px",
              borderRadius: "4px",
              marginBottom: "20px",
              border: "1px solid #ffeeba",
            }}
          >
            <h3 style={{ margin: "0 0 10px 0" }}>
              Word Add-in Handler Not Installed
            </h3>
            <p>
              To open documents with the Word add-in, you need to install the
              Word Add-in Handler application.
            </p>
            <button
              onClick={this.downloadHandlerInstaller}
              style={{
                padding: "8px 16px",
                backgroundColor: "#007bff",
                color: "white",
                border: "none",
                borderRadius: "4px",
                cursor: "pointer",
                fontSize: "14px",
                marginTop: "10px",
              }}
            >
              Download Installer
            </button>
            <p style={{ fontSize: "12px", marginTop: "10px" }}>
              After installation is complete,{" "}
              <a
                href="#"
                onClick={(e) => {
                  e.preventDefault();
                  this.checkAddinHandlerInstalled();
                }}
              >
                click here
              </a>{" "}
              to check again.
            </p>
          </div>
        )}

        {this.state.checkingHandler && (
          <div
            style={{
              padding: "15px",
              backgroundColor: "#e9ecef",
              borderRadius: "4px",
              marginBottom: "20px",
            }}
          >
            Checking if Word Add-in Handler is installed...
          </div>
        )}

        {/* Status Message with Add-in Button */}
        {this.state.status && (
          <div
            className="status-message"
            style={{
              color: this.state.status.includes("Error") ? "red" : "green",
              marginBottom: "20px",
              padding: "10px",
              backgroundColor: "#f8f8f8",
              borderRadius: "4px",
            }}
          >
            {this.state.status}

            {/* Show the Add-in setup button only when needed */}
            {this.state.pendingDocumentUri &&
              this.state.status.includes("Click the button below") && (
                <div style={{ marginTop: "15px" }}>
                  <button
                    onClick={this.handleLaunchAddin}
                    style={{
                      padding: "8px 16px",
                      backgroundColor: "#4CAF50",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      cursor: "pointer",
                      fontSize: "14px",
                    }}
                  >
                    Set Up Add-in for This Document
                  </button>
                </div>
              )}
          </div>
        )}

        {/* Documents Grid */}
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

export default WordAddIn;
