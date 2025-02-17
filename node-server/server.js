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

// Check if Office Word is installed
app.get("/api/check-word", (req, res) => {
  const wordPath =
    "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE";
  const wordPath2 =
    "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.EXE";

  if (fs.existsSync(wordPath) || fs.existsSync(wordPath2)) {
    res.json({ installed: true });
  } else {
    res.json({ installed: false });
  }
});

// Check if add-in is installed
app.get("/api/check-addin", async (req, res) => {
  try {
    const checkCommand =
      'reg query "HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer" /v "UserDevManifests"';

    try {
      await execAsync(checkCommand);
      console.log("Add-in is already installed");
      res.json({ installed: true });
    } catch (error) {
      console.log("Add-in not installed");
      res.json({
        installed: false,
        needsInstallation: true,
        message:
          "Add-in needs to be installed. Would you like to install it now?",
      });
    }
  } catch (error) {
    console.error("Error checking add-in:", error);
    res.status(500).json({
      installed: false,
      error: "Error checking add-in status",
      details: error.message,
    });
  }
});

// Install add-in when user accepts
app.post("/api/install-addin", async (req, res) => {
  try {
    console.log("Starting add-in installation...");

    // Get absolute path to manifest
    const manifestPath = path
      .resolve(__dirname, "../word-add-in/manifest.xml")
      .replace(/\\/g, "\\\\");

    // Create PowerShell installation script
    const installScript = `
      $ManifestPath = "${manifestPath}"
      
      # Kiểm tra manifest
      if (-not (Test-Path $ManifestPath)) {
        throw "Manifest file not found at: $ManifestPath"
      }

      try {
        # Registry paths
        $DevPath = "HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer"
        $TrustedPath = "HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs"
        $WordOptionsPath = "HKCU:\\Software\\Microsoft\\Office\\16.0\\Word\\Options"
        $AddinsPath = "HKCU:\\Software\\Microsoft\\Office\\Word\\Addins"
        
        # Xóa đăng ký cũ
        Remove-Item -Path $DevPath -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "$TrustedPath\\MyAddin" -Recurse -Force -ErrorAction SilentlyContinue
        
        # Tạo registry keys
        New-Item -Path $DevPath -Force | Out-Null
        New-Item -Path "$TrustedPath\\MyAddin" -Force | Out-Null
        New-Item -Path $WordOptionsPath -Force | Out-Null
        
        # Đăng ký manifest trong Developer
        Set-ItemProperty -Path $DevPath -Name "UserDevManifests" -Value $ManifestPath -Type String
        
        # Cấu hình Trusted Catalog
        Set-ItemProperty -Path "$TrustedPath\\MyAddin" -Name "Id" -Value "MyAddin" -Type String
        Set-ItemProperty -Path "$TrustedPath\\MyAddin" -Name "Path" -Value $ManifestPath -Type String
        Set-ItemProperty -Path "$TrustedPath\\MyAddin" -Name "Type" -Value 2 -Type DWord
        Set-ItemProperty -Path "$TrustedPath\\MyAddin" -Name "Flags" -Value 1 -Type DWord
        
        # Bật developer tools
        Set-ItemProperty -Path $WordOptionsPath -Name "DeveloperTools" -Value 1 -Type DWord
        Set-ItemProperty -Path $WordOptionsPath -Name "EnableRibbonCustomization" -Value 1 -Type DWord
        
        # Đăng ký add-in
        $AddinID = [System.IO.Path]::GetFileNameWithoutExtension($ManifestPath)
        $AddinKey = "$AddinsPath\\$AddinID"
        New-Item -Path $AddinKey -Force | Out-Null
        Set-ItemProperty -Path $AddinKey -Name "Description" -Value "My Word Add-in" -Type String
        Set-ItemProperty -Path $AddinKey -Name "FriendlyName" -Value "My Word Add-in" -Type String
        Set-ItemProperty -Path $AddinKey -Name "LoadBehavior" -Value 3 -Type DWord
        Set-ItemProperty -Path $AddinKey -Name "Manifest" -Value $ManifestPath -Type String
        
        # Trust localhost
        CheckNetIsolation.exe LoopbackExempt -a -n="Microsoft.Win32WebViewHost_cw5n1h2txyewy"
        
        # Clear cache
        $WefPath = "$env:LOCALAPPDATA\\Microsoft\\Office\\16.0\\Wef"
        if (Test-Path $WefPath) {
            Remove-Item -Path "$WefPath\\*" -Recurse -Force
        }
        
        # Xóa cache IE
        RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255
        
        # Dừng Word nếu đang chạy
        Get-Process "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 2

        Write-Host "✅ Add-in installed successfully at: $ManifestPath"
        
      } catch {
        Write-Error "Installation error: $_"
        throw
      }
    `;

    // Save and execute script with admin privileges
    const scriptPath = path.join(__dirname, "install-addin.ps1");
    fs.writeFileSync(scriptPath, installScript);

    try {
      await new Promise((resolve, reject) => {
        exec(
          `powershell -ExecutionPolicy Bypass -NoProfile -Command "Start-Process powershell -Verb RunAs -ArgumentList '-ExecutionPolicy Bypass -File \\"${scriptPath}\\"'"`,
          (error, stdout, stderr) => {
            fs.unlinkSync(scriptPath);
            if (error) {
              console.error("Installation error:", error);
              reject(error);
            } else {
              console.log("Installation output:", stdout);
              resolve(stdout);
            }
          }
        );
      });

      console.log("Add-in installation completed successfully");
      res.json({
        installed: true,
        justInstalled: true,
        message:
          "Add-in installed successfully. Please:\n" +
          "1. Ensure dev server is running (npm start)\n" +
          "2. Close all Word instances\n" +
          "3. Clear these caches:\n" +
          "   - %LOCALAPPDATA%\\Microsoft\\Office\\16.0\\Wef\n" +
          "   - Internet Explorer cache\n" +
          "4. Start Word\n" +
          "5. Check Insert > My Add-ins\n" +
          "6. If issues persist:\n" +
          "   - Check manifest.xml is valid\n" +
          "   - Verify localhost:3000 is accessible\n" +
          "   - Run as administrator",
        manifestPath: manifestPath,
      });
    } catch (installError) {
      console.error("Installation failed:", installError);
      res.json({
        installed: false,
        error: "Installation failed. Please check administrator privileges.",
        details: installError.message,
        manifestPath: manifestPath,
      });
    }
  } catch (error) {
    console.error("Server error:", error);
    res.status(500).json({
      installed: false,
      error: "Server error installing add-in",
      details: error.message,
    });
  }
});

// Add verification endpoint
app.get("/api/verify-installation", (req, res) => {
  const verifyScript = `
    $RegistryPath = "HKCU:\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer"
    if (Test-Path $RegistryPath) {
      $manifest = Get-ItemProperty -Path $RegistryPath -Name "UserDevManifests" -ErrorAction SilentlyContinue
      if ($manifest) {
        Write-Host "Add-in is installed"
        exit 0
      }
    }
    Write-Host "Add-in is not installed"
    exit 1
  `;

  exec(`powershell -Command "${verifyScript}"`, (error, stdout) => {
    res.json({
      installed: !error,
      details: stdout.trim(),
    });
  });
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

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
