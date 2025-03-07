require("dotenv").config();

const express = require("express");
const cors = require("cors");
const path = require("path");
const fs = require("fs");
const sudo = require("sudo-prompt");
const app = express();
const port = process.env.REACT_APP_NODE_SERVER_PORT; // Different from your React dev server port
const Registry = require("winreg");
const http = require("http");
const server = http.createServer(app);
const { Server } = require("socket.io");
const io = new Server(server, {
  cors: {
    origin: `http://${process.env.REACT_APP_HOST}:${process.env.REACT_APP_PORT}`, // Your React app URL
    methods: ["GET", "POST"],
    credentials: true,
  },
});
const multer = require("multer");

// Configure CORS to allow requests from all origins
app.use(
  cors({
    origin: function (origin, callback) {
      // Allow any origin
      callback(null, true);
    },
    methods: ["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allowedHeaders: [
      "Content-Type",
      "Accept",
      "Authorization",
      "X-Requested-With",
    ],
    credentials: true,
  })
);

app.use(express.json()); // Parse JSON bodies

// Socket.io connection handling
io.on("connection", (socket) => {
  console.log("Client connected:", socket.id);

  socket.on("disconnect", () => {
    console.log("Client disconnected:", socket.id);
  });
});

// Store document updates
let documentUpdates = [];

// Serve Word documents from a specific directory
app.use("/documents", express.static(path.join(__dirname, "documents")));

// Serve static files
app.use(express.static(path.join(__dirname, "public")));

// Create a downloads directory if it doesn't exist
const downloadsDir = path.join(__dirname, "public", "downloads");
if (!fs.existsSync(downloadsDir)) {
  fs.mkdirSync(downloadsDir, { recursive: true });
}

// Place your WordAddinHandlerSetup.exe in the downloads directory
// You can copy it during server startup if needed:
// fs.copyFileSync(path.join(__dirname, '../path/to/installer/WordAddinHandlerSetup.exe'),
//                 path.join(downloadsDir, 'WordAddinHandlerSetup.exe'));

// Make sure the public directory is served
app.use(
  "/downloads",
  express.static(path.join(__dirname, "public", "downloads"))
);

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
        url: `http://${process.env.REACT_APP_HOST}:${process.env.REACT_APP_NODE_SERVER_PORT}/documents/${file}`,
      }));
    res.json(files);
  } catch (error) {
    res.status(500).json({ error: "Failed to read documents" });
  }
});

// Modified endpoint to receive document updates
app.post("/api/document-update", (req, res) => {
  const { timestamp, documentName, contentLength, eventType } = req.body;

  const update = {
    timestamp,
    documentName,
    contentLength,
    eventType,
    id: Date.now(),
  };

  documentUpdates.push(update);
  console.log("Document updated:", update);

  // Broadcast the update to all connected clients
  io.emit("document-update", update);

  res.json({ message: "Update received", update });
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

// Configure multer for file uploads with limits and proper handling
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    const documentsDir = path.join(__dirname, "documents");
    if (!fs.existsSync(documentsDir)) {
      fs.mkdirSync(documentsDir, { recursive: true });
    }
    cb(null, documentsDir);
  },
  filename: function (req, file, cb) {
    // Generate a temporary filename first to avoid overwriting during upload
    const tempFilename = `temp_${Date.now()}_${file.originalname}`;
    cb(null, tempFilename);
  },
});

// Configure multer with file size limits and error handling
const upload = multer({
  storage: storage,
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB max file size
  },
});

// Endpoint to receive document uploads with metadata
app.post(
  "/api/upload-document-with-metadata",
  upload.single("document"),
  (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({ error: "No file uploaded" });
      }

      // Parse the metadata
      const metadata = req.body.metadata ? JSON.parse(req.body.metadata) : {};

      // Get temporary file path
      const tempFilePath = req.file.path;
      // Final filename (without temp prefix)
      const finalFilename = req.file.originalname || "document.docx";
      // Final file path
      const finalFilePath = path.join(
        path.dirname(tempFilePath),
        finalFilename
      );

      // Rename the file to its original name once upload is complete
      fs.rename(tempFilePath, finalFilePath, (err) => {
        if (err) {
          console.error("Error renaming file:", err);
          return res.status(500).json({
            error: "Failed to save document",
            details: err.message,
          });
        }

        console.log(`Document uploaded and saved: ${finalFilePath}`);
        console.log("Document metadata:", metadata);

        // Add the file info to document updates
        const updateRecord = {
          ...metadata,
          id: Date.now(),
          filePath: finalFilePath,
          fileName: finalFilename,
        };

        documentUpdates.push(updateRecord);

        // Broadcast update notification to all connected clients
        io.emit("document-update", updateRecord);

        res.json({
          message: "Document and metadata uploaded successfully",
          fileName: finalFilename,
          metadata: metadata,
        });
      });
    } catch (error) {
      console.error("Error saving uploaded document:", error);
      // Try to clean up temp file if it exists
      if (req.file && req.file.path) {
        try {
          fs.unlinkSync(req.file.path);
        } catch (cleanupError) {
          console.error("Error cleaning up temp file:", cleanupError);
        }
      }
      res.status(500).json({
        error: "Failed to process document",
        details: error.message,
      });
    }
  }
);

// Change from app.listen to server.listen
server.listen(port, () => {
  console.log(`Server running at http://${process.env.REACT_APP_HOST}:${port}`);
});
