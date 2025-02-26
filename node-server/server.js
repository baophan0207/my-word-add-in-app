const express = require("express");
const cors = require("cors");
const { exec } = require("child_process");
const path = require("path");
const fs = require("fs");
const util = require("util");
const sudo = require("sudo-prompt");
const execAsync = util.promisify(exec);
const app = express();
const port = 3001; // Different from your React dev server port
const Registry = require("winreg");

app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header(
    "Access-Control-Allow-Headers",
    "Origin, X-Requested-With, Content-Type, Accept"
  );
  next();
});

app.use(cors()); // Enable CORS for all routes
app.use(express.json()); // Parse JSON bodies

// Store document updates
let documentUpdates = [];

// Serve Word documents from a specific directory
app.use("/documents", express.static(path.join(__dirname, "documents")));

// Serve static files
app.use(express.static(path.join(__dirname, "public")));

// Get list of available documents
app.get("/api/documents", (req, res) => {
  const documentsPath = path.join(__dirname, "documents");
  try {
    const files = fs
      .readdirSync(documentsPath)
      .filter((file) => file.endsWith(".docx"))
      .map((file) => ({
        id: file,
        name: file,
        url: `http://localhost:3001/documents/${file}`,
      }));
    res.json(files);
  } catch (error) {
    res.status(500).json({ error: "Failed to read documents" });
  }
});

// Endpoint to receive document updates
app.post("/api/document-update", (req, res) => {
  const { timestamp, previousLength, currentLength } = req.body;

  const update = {
    timestamp,
    previousLength,
    currentLength,
    id: Date.now(),
  };

  documentUpdates.push(update);
  console.log("Document updated:", update);

  res.json({ message: "Update received", update });
});

// Endpoint to get all updates
app.get("/api/document-updates", (req, res) => {
  res.json(documentUpdates);
});

// 1. Check if Microsoft Word is installed
app.get("/api/check-word", async (req, res) => {
  try {
    const regKey = new Registry({
      hive: Registry.HKLM,
      key: "\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\WINWORD.EXE",
    });

    regKey.keyExists((err, exists) => {
      if (err) {
        res.status(500).json({
          error: "Error checking Word installation",
          details: err.message,
        });
        return;
      }
      res.json({ isWordInstalled: exists });
    });
  } catch (error) {
    res.status(500).json({
      error: "Failed to check Word installation",
      details: error.message,
    });
  }
});

// Setup Office Add-in
app.post("/api/setup-office-addin", async (req, res) => {
  try {
    const scriptPath = path.join(
      __dirname,
      "../word-add-in/Setup-OfficeAddin.ps1"
    );

    // Verify script exists
    if (!fs.existsSync(scriptPath)) {
      res.status(500).json({
        error: "Setup script not found",
        details: "The PowerShell setup script could not be found",
      });
      return;
    }

    const options = {
      name: "WordAddinSetup",
    };

    // Run PowerShell script with elevated privileges
    sudo.exec(
      `powershell.exe -ExecutionPolicy Bypass -NoProfile -File "${scriptPath}" -documentUrl "${req.body.documentUrl}"`,
      options,
      (error, stdout, stderr) => {
        if (error) {
          console.error("Setup error:", error);
          res.status(500).json({
            error: "Failed to setup Office Add-in",
            details: error.message,
            stdout: stdout,
            stderr: stderr,
          });
          return;
        }

        res.json({
          success: true,
          message: "Office Add-in setup completed successfully",
          output: stdout,
        });
      }
    );
  } catch (error) {
    res.status(500).json({
      error: "Failed to initiate setup",
      details: error.message,
    });
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
