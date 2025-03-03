const { execFile } = require("child_process");
const url = require("url");
const path = require("path");
const fs = require("fs");
const { spawn } = require("child_process");
const sudo = require("sudo-prompt");

// Log startup and arguments
const logFile = path.join(require("os").homedir(), "wordaddin-log.txt");
fs.appendFileSync(
  logFile,
  `\n\nHandler started at ${new Date().toString()}\n`,
  "utf8"
);
fs.appendFileSync(
  logFile,
  `Arguments: ${JSON.stringify(process.argv)}\n`,
  "utf8"
);

try {
  // Get URI from command-line argument
  const uri = process.argv[2]; // e.g., "wordaddin://setup?documentUrl=http://path/to/doc.docx"
  fs.appendFileSync(logFile, `Parsing URI: ${uri}\n`, "utf8");

  // Parse the URI
  const parsedUrl = new URL(uri);
  if (parsedUrl.protocol !== "wordaddin:") {
    throw new Error("Invalid protocol");
  }

  // Get document URL parameter
  const documentUrl = parsedUrl.searchParams.get("documentUrl");
  if (!documentUrl) {
    throw new Error("No document URL specified");
  }

  fs.appendFileSync(logFile, `Document URL: ${documentUrl}\n`, "utf8");

  // Get the directory where the EXE is installed
  const exePath = process.argv[0];
  const exeDir = path.dirname(exePath);

  // Path to the PowerShell script - use the installed location
  const scriptPath = path.join(exeDir, "scripts", "Setup-OfficeAddin.ps1");

  fs.appendFileSync(logFile, `Executable Path: ${exePath}\n`, "utf8");
  fs.appendFileSync(logFile, `Executable Directory: ${exeDir}\n`, "utf8");
  fs.appendFileSync(logFile, `Script Path: ${scriptPath}\n`, "utf8");

  // Verify script exists
  if (!fs.existsSync(scriptPath)) {
    fs.appendFileSync(
      logFile,
      `ERROR: Script not found at: ${scriptPath}\n`,
      "utf8"
    );

    // Try alternative locations for troubleshooting
    const altLocations = [
      path.join(exeDir, "Setup-OfficeAddin.ps1"),
      path.join(__dirname, "scripts", "Setup-OfficeAddin.ps1"),
    ];

    let found = false;
    for (const loc of altLocations) {
      if (fs.existsSync(loc)) {
        fs.appendFileSync(
          logFile,
          `Found script at alternative location: ${loc}\n`,
          "utf8"
        );
        found = true;
      }
    }

    if (!found) {
      throw new Error(`Script not found at: ${scriptPath}`);
    }
  }

  fs.appendFileSync(
    logFile,
    `Executing script for document: ${documentUrl}\n`,
    "utf8"
  );

  // Create the full command string
  const fullCommand = `powershell.exe -ExecutionPolicy Bypass -NoProfile -File "${scriptPath}" -documentUrl "${documentUrl}"`;

  fs.appendFileSync(logFile, `Command: ${fullCommand}\n`, "utf8");

  // Options for sudo-prompt
  const options = {
    name: "WordAddin", // Title that appears in the UAC dialog
  };

  // Execute with elevation
  sudo.exec(fullCommand, options, (error, stdout, stderr) => {
    if (error) {
      fs.appendFileSync(
        logFile,
        `ERROR executing with elevation: ${error}\n and name: ${options.name}`,
        "utf8"
      );
    }
    if (stdout) {
      fs.appendFileSync(logFile, `OUTPUT: ${stdout}\n`, "utf8");
    }
    if (stderr) {
      fs.appendFileSync(logFile, `STDERR: ${stderr}\n`, "utf8");
    }
    fs.appendFileSync(logFile, `Elevated execution completed\n`, "utf8");
  });

  fs.appendFileSync(logFile, `Elevation requested!\n`, "utf8");
} catch (error) {
  fs.appendFileSync(
    logFile,
    `ERROR: ${error.message}\n${error.stack}\n`,
    "utf8"
  );
  console.error(`Error: ${error.message}`);
}
