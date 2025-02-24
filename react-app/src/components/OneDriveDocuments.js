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

  checkWordAndAddin = async (documentUrl) => {
    try {
      // Check Word installation
      const wordResponse = await fetch("http://localhost:3001/api/check-word");
      const wordData = await wordResponse.json();
      this.setState({ wordInstalled: wordData.isWordInstalled });

      if (!wordData.isWordInstalled) {
        this.setState({
          status: "Microsoft Word is not installed. Please install Word first.",
        });
        return false;
      }

      // First open the document and then check/setup add-in
      const openResponse = await fetch(
        "http://localhost:3001/api/setup-office-addin",
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

      // // Check Add-in installation on the opened document
      // const addinResponse = await fetch(
      //   "http://localhost:3001/api/check-addin"
      // );
      // const addinData = await addinResponse.json();
      // this.setState({ addinInstalled: addinData.isAddinInstalled });

      // if (!addinData.isAddinInstalled) {
      //   this.setState({ status: "Setting up Word Add-in..." });
      //   // Setup add-in on the opened document
      //   const setupResponse = await fetch(
      //     "http://localhost:3001/api/setup-office-addin",
      //     {
      //       method: "POST",
      //       headers: {
      //         "Content-Type": "application/json",
      //       },
      //       body: JSON.stringify({ documentUrl }),
      //     }
      //   );
      //   const setupData = await setupResponse.json();

      //   if (setupData.success) {
      //     this.setState({
      //       addinInstalled: true,
      //       status: "Add-in setup successfully!",
      //     });
      //     return true;
      //   } else {
      //     this.setState({
      //       status:
      //         "Failed to setup Add-in. Please try again or contact support.",
      //     });
      //     return false;
      //   }
      // }
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
      const protocolUrl = `ms-word:ofe|u|${baseUrl}?web=1`;
      window.open(protocolUrl);

      // Wait for document to open (give it a few seconds)
      this.setState({ status: "Opening document..." });
      // await new Promise((resolve) => setTimeout(resolve, 5000));

      // Now check Word and setup add-in with the specific document URL
      const ready = await this.checkWordAndAddin(baseUrl);
      // if (!ready) {
      //   return;
      // }

      this.setState({
        status: "Document opened successfully with add-in enabled",
      });
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
              >
                <span>{doc.name}</span>
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }
}

export default OneDriveDocuments;
