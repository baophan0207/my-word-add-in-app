[Setup]
AppName=Word Add-in Handler
AppVersion=1.0
DefaultDirName={commonpf}\WordAddinHandler
DefaultGroupName=Word Add-in Handler
UninstallDisplayIcon={app}\word-addin-handler.exe
Compression=lzma2
SolidCompression=yes
OutputDir=.
PrivilegesRequired=admin

[Files]
; Copy the executable
Source: "word-addin-handler.exe"; DestDir: "{app}"; Flags: ignoreversion

; Copy the PowerShell script and any other required files
Source: "scripts\*"; DestDir: "{app}\scripts"; Flags: ignoreversion recursesubdirs

; Include the register-uri-handler script
Source: "register-uri-handler.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Run]
; Register the URI handler automatically during installation
Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -WindowStyle Hidden -File ""{app}\register-uri-handler.ps1"" -ExePath ""{app}\word-addin-handler.exe"""; Flags: runhidden

; Set PowerShell execution policy to allow the script to run
Filename: "powershell.exe"; Parameters: "-Command ""Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force"""; Flags: runhidden

[UninstallRun]
; Remove the URI protocol registration during uninstallation
Filename: "powershell.exe"; Parameters: "-Command ""Remove-Item -Path 'HKCU:\Software\Classes\wordaddin' -Recurse -Force -ErrorAction SilentlyContinue"""; Flags: runhidden

[Icons]
Name: "{group}\Uninstall Word Add-in Handler"; Filename: "{uninstallexe}" 