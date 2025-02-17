import React from "react";
import "./OneDriveDocument.css";

class OneDriveDocuments extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      documents: [],
      wordInstalled: false,
      addinInstalled: false,
      status: "",
    };
  }

  componentDidMount() {
    this.loadDocuments();
    this.checkWordAndAddin();
  }

  loadDocuments = async () => {
    try {
      const response = await fetch("http://localhost:3001/api/documents");
      const documents = await response.json();
      this.setState({ documents });
    } catch (error) {
      console.error("Error loading documents:", error);
      this.setState({ status: "Error loading documents" });
    }
  };

  checkWordAndAddin = async () => {
    try {
      // Check Word installation
      const wordResponse = await fetch("http://localhost:3001/api/check-word");
      const wordData = await wordResponse.json();

      console.log("Word data:", wordData);

      this.setState({
        wordInstalled: wordData.installed,
        status: wordData.installed ? "" : "Microsoft Word is not installed",
      });

      if (!wordData.installed) return;

      // Check Add-in installation
      const addinResponse = await fetch(
        "http://localhost:3001/api/check-addin"
      );
      const addinData = await addinResponse.json();

      console.log("Addin data:", addinData);

      if (addinData.needsInstallation) {
        // Ask user for confirmation
        if (window.confirm(addinData.message)) {
          // User accepted, proceed with installation
          const installResponse = await fetch(
            "http://localhost:3001/api/install-addin",
            {
              method: "POST",
              headers: {
                "Content-Type": "application/json",
              },
            }
          );
          const installData = await installResponse.json();

          if (installData.error) {
            this.setState({
              status: installData.error,
              addinInstalled: false,
            });
          } else {
            this.setState({
              addinInstalled: true,
              status: "Add-in installed successfully. Please restart Word.",
            });
          }
        } else {
          // User declined installation
          this.setState({
            status: "Add-in installation was declined",
            addinInstalled: false,
          });
        }
      } else if (addinData.installed) {
        this.setState({
          addinInstalled: true,
          status: "",
        });
      }
    } catch (error) {
      console.error("Error checking Word and Add-in:", error);
      this.setState({
        status: "Error checking installations",
        addinInstalled: false,
      });
    }
  };

  openDocument = async (doc) => {
    try {
      if (!this.state.wordInstalled) {
        this.setState({ status: "Please install Microsoft Word first" });
        return;
      }

      if (!this.state.addinInstalled) {
        await this.checkWordAndAddin(); // Try to install add-in
        if (!this.state.addinInstalled) {
          this.setState({
            status: "Please wait while the add-in is being installed",
          });
          return;
        }
      }

      // Use the document URL from the server
      const baseUrl = doc.url;

      // Add parameters for add-in
      const addInId = "a8b28819-f6c0-42f7-b7c3-460fd297efa4"; // Your add-in ID
      const addInVersion = "1.0.0.0";
      const addInUrl = "https://localhost:3000/taskpane.html"; // Your add-in URL

      // Create protocol handler URL with add-in parameters
      const protocolUrl = `ms-word:ofe|u|${baseUrl}?web=1&wdaddinId=${addInId}&wdAddinVersion=${addInVersion}&wdAddinUrl=${encodeURIComponent(
        addInUrl
      )}`;

      // Open Word Desktop
      window.open(protocolUrl);
    } catch (error) {
      console.error("Error opening document:", error);
      this.setState({ status: "Error opening document" });
    }
  };

  render() {
    return (
      <div className="documents-list">
        <h2>Your Word Documents</h2>
        {this.state.status && (
          <div
            className="status-message"
            style={{
              color: this.state.status.includes("success") ? "green" : "red",
            }}
          >
            {this.state.status}
          </div>
        )}
        <div className="document-container">
          <div className="documents-grid">
            {this.state.documents.map((doc) => (
              <div
                key={doc.id}
                className="document-item"
                onClick={() => this.openDocument(doc)}
                style={{
                  opacity:
                    !this.state.wordInstalled || !this.state.addinInstalled
                      ? 0.5
                      : 1,
                  cursor:
                    !this.state.wordInstalled || !this.state.addinInstalled
                      ? "not-allowed"
                      : "pointer",
                }}
              >
                <span>{doc.name}</span>
                {(!this.state.wordInstalled || !this.state.addinInstalled) && (
                  <span className="installation-required">
                    {!this.state.wordInstalled
                      ? "Word Required"
                      : "Installing Add-in..."}
                  </span>
                )}
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }
}

export default OneDriveDocuments;
