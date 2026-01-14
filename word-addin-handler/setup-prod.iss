[Setup]
AppName=IP Agent Word Handler
AppVersion=1.0
DefaultDirName={commonpf}\IPAgentWordHandler
DefaultGroupName=IP Agent Word Handler
UninstallDisplayIcon={app}\IPAgentWordHandler.exe
Compression=lzma2
SolidCompression=yes
OutputDir=.
OutputBaseFilename=WordAddinHandlerSetup
PrivilegesRequired=admin
AlwaysRestart=no
; The following directives ensure the application installs for all users
AllowUNCPath=false
AlwaysUsePersonalGroup=false
; Always create a fresh installation directory
DirExistsWarning=auto

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

; Set PowerShell execution policy to allow the script to run for all users
Filename: "powershell.exe"; Parameters: "-Command ""Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned -Force"""; Flags: runhidden

[UninstallRun]
; Stop any running processes before uninstallation
Filename: "taskkill"; Parameters: "/F /IM word-addin-handler.exe /T"; Flags: runhidden; RunOnceId: "StopHandler"

; Remove the URI protocol registration during uninstallation
Filename: "powershell.exe"; Parameters: "-Command ""Remove-Item -Path 'HKLM:\Software\Classes\wordaddin' -Recurse -Force -ErrorAction SilentlyContinue"""; Flags: runhidden; RunOnceId: "RemoveRegistry"

; Force remove any remaining files and folders
Filename: "powershell.exe"; Parameters: "-Command ""Start-Sleep -Seconds 2; if (Test-Path '{app}') {{ Remove-Item -Path '{app}' -Recurse -Force -ErrorAction SilentlyContinue }}"""; Flags: runhidden; RunOnceId: "ForceCleanup"

[UninstallDelete]
; Delete all files in the installation directory
Type: files; Name: "{app}\*.*"
Type: filesandordirs; Name: "{app}\scripts"
Type: filesandordirs; Name: "{app}\logs"
Type: filesandordirs; Name: "{app}\temp"
; Delete the main installation directory
Type: dirifempty; Name: "{app}"

[Code]
// Custom uninstall procedure for thorough cleanup
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  AppDir: String;
  ResultCode: Integer;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    AppDir := ExpandConstant('{app}');
    
    // Force kill any remaining processes
    Exec('taskkill', '/F /IM word-addin-handler.exe /T', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    
    // Wait a moment for processes to terminate
    Sleep(1000);
    
    // Use PowerShell to force remove the directory
    Exec('powershell.exe', 
         '-Command "if (Test-Path ''' + AppDir + ''') { Remove-Item -Path ''' + AppDir + ''' -Recurse -Force -ErrorAction SilentlyContinue }"', 
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    
    // Also clean up any temp files that might be left
    Exec('powershell.exe', 
         '-Command "Get-ChildItem -Path $env:TEMP -Filter ''*word-addin*'' -Recurse | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue"', 
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;

// Pre-installation cleanup to handle existing installations
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  AppDir: String;
  ResultCode: Integer;
begin
  AppDir := ExpandConstant('{app}');
  
  // If directory exists, clean it up before installation
  if DirExists(AppDir) then
  begin
    // Stop any running processes
    Exec('taskkill', '/F /IM word-addin-handler.exe /T', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Sleep(1000);
    
    // Try to remove existing installation
    Exec('powershell.exe', 
         '-Command "if (Test-Path ''' + AppDir + ''') { Remove-Item -Path ''' + AppDir + ''' -Recurse -Force -ErrorAction SilentlyContinue }"', 
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
  
  Result := '';
end;

[Icons]
Name: "{group}\Uninstall IP Agent Word Handler"; Filename: "{uninstallexe}"
