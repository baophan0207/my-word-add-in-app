const path = require("path");
const fs = require("fs");
const sudo = require("sudo-prompt");
const { io } = require("socket.io-client");

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

// Socket.io connection setup
let socket = null;

function connectToServer() {
  try {
    // Connect to the socket.io server
    socket = io("https://test.ipagent.ai", {
      autoConnect: true,
      reconnection: true,
      reconnectionDelay: 1000,
      reconnectionAttempts: 5,
    });

    socket.on("connect", () => {
      fs.appendFileSync(
        logFile,
        `Connected to socket.io server: ${socket.id}\n`,
        "utf8"
      );

      // Register as a handler client
      socket.emit("register-client", {
        clientType: "handler",
        clientInfo: {
          version: "1.0.0",
          processId: process.pid,
          startTime: new Date(),
          platform: process.platform,
        },
      });

      // Process request after socket is connected
      fs.appendFileSync(
        logFile,
        `Socket connected, now processing request...\n`,
        "utf8"
      );
      setTimeout(processRequest, 500); // Small delay to ensure connection is stable
    });

    socket.on("connect_error", (error) => {
      fs.appendFileSync(
        logFile,
        `Socket connection error: ${error.message}\n`,
        "utf8"
      );
    });

    socket.on("disconnect", (reason) => {
      fs.appendFileSync(logFile, `Socket disconnected: ${reason}\n`, "utf8");
    });
  } catch (error) {
    fs.appendFileSync(
      logFile,
      `Error setting up socket connection: ${error.message}\n`,
      "utf8"
    );
  }
}

// Send single protocol response to server with session and document info
function sendProtocolResponse(
  success,
  documentName,
  sessionId,
  message = null
) {
  const response = {
    success,
    documentName,
    sessionId,
    timestamp: new Date(),
    processId: process.pid,
    message:
      message ||
      (success
        ? "Handler processed request successfully"
        : "Handler failed to process request"),
  };

  if (socket && socket.connected) {
    socket.emit("protocol-response", response);
    fs.appendFileSync(
      logFile,
      `Protocol response sent: ${JSON.stringify(response)}\n`,
      "utf8"
    );
  } else {
    fs.appendFileSync(
      logFile,
      `WARNING: Socket not connected, cannot send response: ${JSON.stringify(
        response
      )}\n`,
      "utf8"
    );
  }
}

// Wait for socket connection and then process the request
function processRequest() {
  try {
    // Get URI from command-line argument
    const uri = process.argv[2];
    fs.appendFileSync(logFile, `Parsing URI: ${uri}\n`, "utf8");

    // Parse the URI
    const parsedUrl = new URL(uri);
    if (parsedUrl.protocol !== "wordaddin:") {
      throw new Error("Invalid protocol");
    }

    // Generate or get sessionId from URL parameters
    const sessionId = parsedUrl.searchParams.get("sessionId");

    // Handle ping request - special case for testing protocol
    if (parsedUrl.pathname === "/ping/" || parsedUrl.pathname === "/ping") {
      fs.appendFileSync(
        logFile,
        `Ping request received - Protocol working correctly\n`,
        "utf8"
      );
      console.log("WordAddin protocol handler is working correctly");

      // Send successful protocol response
      sendProtocolResponse(true, "ping", sessionId, "Protocol test successful");

      // Exit after sending response
      setTimeout(() => {
        if (socket) socket.disconnect();
        process.exit(0);
      }, 2000);
      return;
    }

    // Get parameters from URL for setup requests
    const documentName = parsedUrl.searchParams.get("documentName");
    const documentUrl = parsedUrl.searchParams.get("documentUrl") || "";

    if (!documentName) {
      throw new Error("No document name specified");
    }

    // Send success response immediately - handler is available and can process the request
    sendProtocolResponse(
      true,
      documentName,
      sessionId,
      "Word Add-in Handler is available and processing document"
    );

    fs.appendFileSync(logFile, `Session ID: ${sessionId}\n`, "utf8");
    fs.appendFileSync(logFile, `Document Name: ${documentName}\n`, "utf8");
    fs.appendFileSync(logFile, `Document URL: ${documentUrl}\n`, "utf8");

    // Get the directory where the EXE is installed
    const exePath = process.argv[0];
    const exeDir = path.dirname(exePath);

    // Path to the PowerShell script - use the installed location
    const scriptPath = path.join(exeDir, "scripts", "Setup-OfficeAddin.ps1");

    fs.appendFileSync(logFile, `Script Path: ${scriptPath}\n`, "utf8");

    // Verify script exists
    if (!fs.existsSync(scriptPath)) {
      fs.appendFileSync(
        logFile,
        `ERROR: Script not found at: ${scriptPath}\n`,
        "utf8"
      );

      sendProtocolResponse(
        false,
        documentName,
        sessionId,
        `Script not found at: ${scriptPath}`
      );

      setTimeout(() => {
        if (socket) socket.disconnect();
        process.exit(1);
      }, 2000);
      return;
    }

    fs.appendFileSync(
      logFile,
      `Handler available - sent success response. Now executing script for document: ${documentName}${
        documentUrl ? ` with URL: ${documentUrl}` : ""
      }\n`,
      "utf8"
    );

    // Create the full command string with both parameters
    let fullCommand = `powershell -ExecutionPolicy Bypass -File "${scriptPath}" -documentName "${documentName}"`;

    // Add documentUrl parameter if provided
    if (documentUrl) {
      fullCommand += ` -documentUrl "${documentUrl}"`;
    }

    fs.appendFileSync(logFile, `Command: ${fullCommand}\n`, "utf8");

    // Options for sudo-prompt
    const options = {
      name: "WordAddin",
    };

    fs.appendFileSync(
      logFile,
      `Requesting elevation for script execution...\n`,
      "utf8"
    );

    // Execute with elevation - but don't send response, just log the result
    sudo.exec(fullCommand, options, (error, stdout, stderr) => {
      if (error) {
        fs.appendFileSync(
          logFile,
          `ERROR executing with elevation: ${error}\n`,
          "utf8"
        );
        fs.appendFileSync(
          logFile,
          `Script execution failed, but success response was already sent\n`,
          "utf8"
        );
      } else {
        if (stdout) {
          fs.appendFileSync(logFile, `OUTPUT: ${stdout}\n`, "utf8");
        }
        if (stderr) {
          fs.appendFileSync(logFile, `STDERR: ${stderr}\n`, "utf8");
        }
        fs.appendFileSync(
          logFile,
          `Elevated execution completed successfully\n`,
          "utf8"
        );
      }

      // Disconnect after completion (whether success or failure)
      setTimeout(() => {
        if (socket) socket.disconnect();
        process.exit(0); // Always exit with success since we already sent response
      }, 2000);
    });

    fs.appendFileSync(
      logFile,
      `Elevation requested! Handler response already sent.\n`,
      "utf8"
    );
  } catch (error) {
    fs.appendFileSync(
      logFile,
      `ERROR: ${error.message}\n${error.stack}\n`,
      "utf8"
    );
    console.error(`Error: ${error.message}`);

    // Send failed protocol response
    sendProtocolResponse(false, error.message, "unknown", error.message);

    // Disconnect on error
    setTimeout(() => {
      if (socket) socket.disconnect();
      process.exit(1);
    }, 2000);
  }
}

try {
  // Initialize socket connection
  connectToServer();

  // Timeout if connection takes too long
  setTimeout(() => {
    if (!socket || !socket.connected) {
      fs.appendFileSync(
        logFile,
        `Connection timeout - processing request without server connection\n`,
        "utf8"
      );
      processRequest(); // Try to process anyway, but response won't be sent
    }
  }, 5000); // 5 second timeout
} catch (error) {
  fs.appendFileSync(
    logFile,
    `FATAL ERROR: ${error.message}\n${error.stack}\n`,
    "utf8"
  );
  console.error(`Fatal error: ${error.message}`);
  process.exit(1);
}
