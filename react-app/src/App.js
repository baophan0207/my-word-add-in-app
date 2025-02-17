import React from "react";
import "./App.css";
import OneDriveDocuments from "./components/OneDriveDocuments";

function App() {
  return (
    <div className="app">
      <header className="app-header">
        <h1>Document Viewer</h1>
      </header>
      <main className="app-main">
        <OneDriveDocuments />
      </main>
    </div>
  );
}

export default App;
