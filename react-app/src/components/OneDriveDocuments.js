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
      this.setState({ wordInstalled: wordData.isWordInstalled });

      if (!wordData.isWordInstalled) {
        this.setState({
          status: "Microsoft Word is not installed. Please install Word first.",
        });
        return;
      }

      // Check Add-in installation
      const addinResponse = await fetch(
        "http://localhost:3001/api/check-addin"
      );
      const addinData = await addinResponse.json();
      this.setState({ addinInstalled: addinData.isAddinInstalled });

      if (!addinData.isAddinInstalled) {
        this.setState({ status: "Installing Word Add-in..." });
        // Try to install the add-in
        // const installResponse = await fetch(
        //   "http://localhost:3001/api/install-addin",
        //   {
        //     method: "POST",
        //   }
        // );
        const installResponse = await fetch(
          "http://localhost:3001/api/setup-office-addin",
          {
            method: "POST",
          }
        );
        const installData = await installResponse.json();

        if (installData.success) {
          this.setState({
            addinInstalled: true,
            status: "Add-in installed successfully!",
          });
        } else {
          this.setState({
            status:
              "Failed to install Add-in. Please try again or contact support.",
          });
        }
      }
    } catch (error) {
      console.error("Error checking Word and Add-in:", error);
      this.setState({ status: "Error checking Word and Add-in installation" });
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

      // Create protocol handler URL with add-in parameters
      const protocolUrl = `ms-word:ofe|u|${baseUrl}?web=1`;

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
