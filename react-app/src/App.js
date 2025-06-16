import React from "react";
import "./App.css";
import WordAddIn from "./components/WordAddIn/WordAddIn";
import DocumentEditor from "./components/DocumentEditor/DocumentEditor";
class App extends React.Component {
  render() {
    return (
      <div className="app">
        <main className="app-main">
          <WordAddIn />
        </main>
        {/* <DocumentEditor /> */}
      </div>
    );
  }
}

export default App;
