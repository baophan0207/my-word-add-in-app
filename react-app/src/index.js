import React from "react";
import ReactDOM from "react-dom";
import App from "./App";
import "./index.css";
import { registerLicense } from "@syncfusion/ej2-base";
registerLicense(
  "Ngo9BigBOggjHTQxAR8/V1NMaF1cXmhNYVppR2Nbek5xdV9GZ1ZVTGY/P1ZhSXxWdkZiWX1ddXZVT2RbUUI="
);

ReactDOM.render(<App />, document.getElementById("root"));
