import React from "react";
import "./App.css";
import OneDriveDocuments from "./components/OneDriveDocument/OneDriveDocuments";
import DocumentEditor from "./components/DocumentEditor/DocumentEditor";
class App extends React.Component {
  render() {
    return (
      <div className="app">
        {/* <main className="app-main">
          <OneDriveDocuments />
        </main> */}

        <DocumentEditor />
      </div>
    );
  }
}

export default App;
