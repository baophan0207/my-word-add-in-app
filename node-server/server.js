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

// 2. Check if Add-in is installed
app.get("/api/check-addin", async (req, res) => {
  try {
    const regKey = new Registry({
      hive: Registry.HKCU,
      key: "\\Software\\Microsoft\\Office\\Word\\Addins\\f85491a7-0cf8-4950-b18c-d85ae9970d61", // Replace with your add-in ID
    });

    regKey.keyExists((err, exists) => {
      if (err) {
        res.status(500).json({
          error: "Error checking Add-in installation",
          details: err.message,
        });
        return;
      }
      res.json({ isAddinInstalled: exists });
    });
  } catch (error) {
    res.status(500).json({
      error: "Failed to check Add-in installation",
      details: error.message,
    });
  }
});

// 3. Install Add-in
app.post("/api/install-addin", async (req, res) => {
  const manifestPath =
    "file:///" +
    path.join(__dirname, "../word-add-in/manifest.xml").replace(/\\/g, "/");
  // Add UTF-16LE BOM and ensure Windows-style line endings
  const regCommand = `\ufeffWindows Registry Editor Version 5.00\r\n\r\n[HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\Word\\Addins\\f85491a7-0cf8-4950-b18c-d85ae9970d61]\r\n"Description"="A template to get started."\r\n"FriendlyName"="word-add-in"\r\n"LoadBehavior"=dword:00000003\r\n"Manifest"="${manifestPath}"\r\n"Type"="Manifest"\r\n`;

  const regFilePath = path.join(__dirname, "install-addin.reg");

  try {
    // Verify manifest file exists (use the actual file path, not the file:/// URL)
    const actualManifestPath = path.join(
      __dirname,
      "../word-add-in/manifest.xml"
    );
    if (!fs.existsSync(actualManifestPath)) {
      res.status(500).json({
        error: "Manifest file not found",
        details:
          "The add-in manifest file could not be found at the specified location",
      });
      return;
    }

    // Write the file with UTF-16LE encoding
    await fs.promises.writeFile(regFilePath, regCommand, {
      encoding: "utf16le",
    });

    const options = {
      name: "WordAddinInstaller",
    };

    sudo.exec(
      `reg import "${regFilePath}"`,
      options,
      (error, stdout, stderr) => {
        // Clean up the temporary registry file
        fs.unlink(regFilePath, (unlinkError) => {
          if (unlinkError) {
            console.error("Error cleaning up registry file:", unlinkError);
          }
        });

        if (error) {
          res.status(500).json({
            error: "Failed to install Add-in",
            details: error.message,
          });
          return;
        }

        res.json({
          success: true,
          message: "Add-in installed successfully",
        });
      }
    );
  } catch (error) {
    res.status(500).json({
      error: "Failed to prepare Add-in installation",
      details: error.message,
    });
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
