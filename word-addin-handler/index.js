const path = require("path");
const fs = require("fs");
const http = require("http");
const sudo = require("sudo-prompt");

// Configuration for local HTTP server (for protocol detection)
const LOCAL_SERVER_PORT = 9876;
const LOCAL_SERVER_TIMEOUT = 5000; // Giảm xuống 5 giây
const SHUTDOWN_GRACE_PERIOD = 500; // 500ms để cleanup

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

// Graceful shutdown function
function gracefulShutdown(server, exitCode = 0) {
  fs.appendFileSync(logFile, `Initiating graceful shutdown...\n`, "utf8");

  if (server) {
    server.close(() => {
      fs.appendFileSync(logFile, `Server closed successfully\n`, "utf8");

      // Force exit after grace period
      setTimeout(() => {
        fs.appendFileSync(
          logFile,
          `Exiting process with code ${exitCode}\n`,
          "utf8"
        );
        process.exit(exitCode);
      }, SHUTDOWN_GRACE_PERIOD);
    });

    // Destroy all active connections to force immediate closure
    server.closeAllConnections();
  } else {
    setTimeout(() => {
      process.exit(exitCode);
    }, SHUTDOWN_GRACE_PERIOD);
  }
}

// Process the request from the custom protocol
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

    // Get parameters from URL
    const documentName = parsedUrl.searchParams.get("documentName");
    const documentUrl = parsedUrl.searchParams.get("documentUrl") || "";

    fs.appendFileSync(logFile, `Pathname: ${parsedUrl.pathname}\n`, "utf8");
    fs.appendFileSync(
      logFile,
      `Document Name: ${documentName || "NOT PROVIDED"}\n`,
      "utf8"
    );
    fs.appendFileSync(
      logFile,
      `Document URL: ${documentUrl || "NOT PROVIDED"}\n`,
      "utf8"
    );

    // ===== Nếu KHÔNG có documentName → Chạy HTTP Server =====
    if (!documentName) {
      fs.appendFileSync(
        logFile,
        `No document name - Starting HTTP server for detection\n`,
        "utf8"
      );

      let requestCount = 0;
      const server = http.createServer((req, res) => {
        requestCount++;

        fs.appendFileSync(
          logFile,
          `[Request #${requestCount}] ${req.method} ${req.url}\n`,
          "utf8"
        );

        // Set CORS headers
        res.setHeader("Access-Control-Allow-Origin", "*");
        res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
        res.setHeader("Access-Control-Allow-Headers", "Content-Type");
        res.setHeader("Connection", "close"); // Force close connection

        if (req.method === "OPTIONS") {
          res.writeHead(204);
          res.end();
          return;
        }

        if (req.url === "/ping" || req.url === "/ping/") {
          const response = {
            success: true,
            message: "WordAddin handler is installed",
            version: "1.0.0",
            timestamp: new Date().toISOString(),
          };

          res.writeHead(200, { "Content-Type": "application/json" });
          res.end(JSON.stringify(response));

          fs.appendFileSync(
            logFile,
            `Ping successful! Handler detected.\n`,
            "utf8"
          );

          // Shutdown ngay sau khi response ping thành công
          setTimeout(() => {
            fs.appendFileSync(
              logFile,
              `Ping served - initiating shutdown\n`,
              "utf8"
            );
            gracefulShutdown(server, 0);
          }, 100); // 100ms để đảm bảo response được gửi
        } else {
          res.writeHead(404);
          res.end("Not Found");
        }
      });

      server.on("error", (err) => {
        fs.appendFileSync(
          logFile,
          `Server error: ${err.code} - ${err.message}\n`,
          "utf8"
        );

        if (err.code === "EADDRINUSE") {
          fs.appendFileSync(
            logFile,
            `Port ${LOCAL_SERVER_PORT} in use - another instance running\n`,
            "utf8"
          );
          process.exit(0);
        } else {
          process.exit(1);
        }
      });

      server.on("listening", () => {
        const addr = server.address();
        fs.appendFileSync(
          logFile,
          `✓ Server listening on ${addr.address}:${addr.port}\n`,
          "utf8"
        );
      });

      // Set timeout for max server lifetime
      const timeoutId = setTimeout(() => {
        fs.appendFileSync(
          logFile,
          `Timeout reached (${LOCAL_SERVER_TIMEOUT}ms) - no ping received\n`,
          "utf8"
        );
        gracefulShutdown(server, 0);
      }, LOCAL_SERVER_TIMEOUT);

      // Cleanup timeout if server shuts down early
      server.on("close", () => {
        clearTimeout(timeoutId);
      });

      server.listen(LOCAL_SERVER_PORT, "127.0.0.1");

      return; // EXIT
    }

    // ===== Nếu CÓ documentName → Chạy Script Setup =====
    fs.appendFileSync(logFile, `Document setup requested\n`, "utf8");

    const exePath = process.argv[0];
    const exeDir = path.dirname(exePath);
    const scriptPath = path.join(exeDir, "scripts", "Setup-OfficeAddin.ps1");

    fs.appendFileSync(logFile, `Script Path: ${scriptPath}\n`, "utf8");

    if (!fs.existsSync(scriptPath)) {
      fs.appendFileSync(
        logFile,
        `ERROR: Script not found at: ${scriptPath}\n`,
        "utf8"
      );
      process.exit(1);
      return;
    }

    let fullCommand = `powershell -ExecutionPolicy Bypass -File "${scriptPath}" -documentName "${documentName}"`;

    if (documentUrl) {
      fullCommand += ` -documentUrl "${documentUrl}"`;
    }

    fs.appendFileSync(logFile, `Command: ${fullCommand}\n`, "utf8");

    const options = {
      name: "WordAddin",
    };

    fs.appendFileSync(logFile, `Requesting elevation...\n`, "utf8");

    sudo.exec(fullCommand, options, (error, stdout, stderr) => {
      if (error) {
        fs.appendFileSync(logFile, `ERROR: ${error}\n`, "utf8");
      } else {
        if (stdout) {
          fs.appendFileSync(logFile, `OUTPUT: ${stdout}\n`, "utf8");
        }
        if (stderr) {
          fs.appendFileSync(logFile, `STDERR: ${stderr}\n`, "utf8");
        }
        fs.appendFileSync(logFile, `Script completed\n`, "utf8");
      }
      process.exit(0);
    });

    fs.appendFileSync(logFile, `Waiting for elevation...\n`, "utf8");
  } catch (error) {
    fs.appendFileSync(
      logFile,
      `ERROR: ${error.message}\n${error.stack}\n`,
      "utf8"
    );
    process.exit(1);
  }
}

// Start processing
try {
  processRequest();
} catch (error) {
  fs.appendFileSync(
    logFile,
    `FATAL: ${error.message}\n${error.stack}\n`,
    "utf8"
  );
  process.exit(1);
}
