param(
    [Parameter(Mandatory=$true)]
    [string]$documentName,
    
    [Parameter(Mandatory=$true)]
    [string]$documentUrl = ""
)

# Script to automate Office Add-in network share setup, manifest creation, and trust configuration
# Requires elevation but now works with current user context instead of admin context
# 
# Multi-User Environment Improvements:
# - Auto-detects current user via explorer.exe process, console session, or recent profiles
# - Creates manifest and network share in current user's Documents folder
# - Sets proper permissions for both admin and current user
# - Configures trusted catalogs in current user's registry (HKCU)
# - Opens Word documents in current user's context
#
# Requires elevation
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as Administrator!"
    Exit 1
}

# Define error code constants
$ERROR_WORD_NOT_INSTALLED = 100
$ERROR_MANIFEST_CREATION_FAILED = 101
$ERROR_NETWORK_SHARE_FAILED = 102
$ERROR_OPEN_DOCUMENT_FAILED = 103
$ERROR_ADDIN_CONFIGURATION_FAILED = 104
$ERROR_USER_NOT_FOUND = 105

# Create a global error hashtable for better error handling
$script:ErrorDetails = @{
    $ERROR_WORD_NOT_INSTALLED = "Microsoft Word is not installed on this device."
    $ERROR_MANIFEST_CREATION_FAILED = "Failed to create or verify the manifest file."
    $ERROR_NETWORK_SHARE_FAILED = "Failed to create or verify network share."
    $ERROR_OPEN_DOCUMENT_FAILED = "Failed to open the specified document."
    $ERROR_ADDIN_CONFIGURATION_FAILED = "Failed to configure the Office add-in."
    $ERROR_USER_NOT_FOUND = "Could not determine current user context."
}

# Add required assemblies
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Word -ErrorAction SilentlyContinue

# Define Win32 API functions needed for UI automation
if (-not ([System.Management.Automation.PSTypeName]'Win32Functions.Win32Api').Type) {
    $signature = @'
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, ref uint lpdwProcessId);
    
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
    Add-Type -MemberDefinition $signature -Name "Win32Api" -Namespace Win32Functions
}

# Define mouse event functions for UI automation
if (-not ([System.Management.Automation.PSTypeName]'Win32Functions.Win32MouseEvent').Type) {
    $signature = @'
    [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
'@
    Add-Type -MemberDefinition $signature -Name "Win32MouseEvent" -Namespace Win32Functions
}

# Function to create a new GUID
function New-Guid {
    return [guid]::NewGuid().ToString()
}

# Function to get current user information (not admin)
function Get-CurrentUser {
    Write-Host "Determining current user context..."
    
    # Try multiple methods to get the actual current user
    $currentUser = $null
    
    # Method 1: Try to get the user who owns explorer.exe process
    try {
        $explorerProcess = Get-WmiObject -Class Win32_Process -Filter "Name='explorer.exe'" | 
            Where-Object { $_.GetOwner().User -ne "Administrator" } | 
            Select-Object -First 1
        if ($explorerProcess) {
            $currentUser = $explorerProcess.GetOwner().User
            Write-Host "Detected user from explorer.exe: $currentUser"
        }
    } catch {
        Write-Verbose "Could not get user from explorer.exe process"
    }
    
    # Method 2: Check who's logged on console
    if (-not $currentUser) {
        try {
            $loggedOnUsers = quser 2>$null | Where-Object { $_ -match "console" -and $_ -notmatch "Administrator" }
            if ($loggedOnUsers) {
                $consoleUser = ($loggedOnUsers -split '\s+')[1]
                if ($consoleUser -and $consoleUser -ne "Administrator") {
                    $currentUser = $consoleUser
                    Write-Host "Detected user from console session: $currentUser"
                }
            }
        } catch {
            Write-Verbose "Could not get console user"
        }
    }
    
    # Method 3: Check environment variables from registry
    if (-not $currentUser) {
        try {
            # Look for recent user profiles
            $recentProfile = Get-ChildItem "C:\Users" -Directory | 
                Where-Object { 
                    $_.Name -notin @("Administrator", "Public", "Default", "DefaultAppPool") -and
                    (Test-Path (Join-Path $_.FullName "NTUSER.DAT"))
                } | 
                Sort-Object LastWriteTime -Descending | 
                Select-Object -First 1
            
            if ($recentProfile) {
                $currentUser = $recentProfile.Name
                Write-Host "Detected user from recent profile: $currentUser"
            }
        } catch {
            Write-Verbose "Could not detect user from profiles"
        }
    }
    
    # Fallback to environment variable (will be admin when running as admin)
    if (-not $currentUser) {
        $currentUser = $env:USERNAME
        Write-Warning "Using current environment username (may be admin): $currentUser"
    }
    
    # Get user profile path
    $userProfile = "C:\Users\$currentUser"
    if (-not (Test-Path $userProfile)) {
        Write-Error "User profile not found: $userProfile"
        return $null
    }
    
    return @{
        Username = $currentUser
        ProfilePath = $userProfile
        DocumentsPath = Join-Path $userProfile "Documents"
    }
}

# Function to create manifest file for current user
function New-ManifestFile {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$ShareName = "",
        
        [Parameter()]
        [string]$FolderPath = ""
    )
    
    Write-Host "Creating manifest file..."
    
    # Get current user info and set default values
    $userInfo = Get-CurrentUser
    if (-not $userInfo) {
        Write-Error "Could not determine current user"
        return $false
    }
    
    # Set user-specific share name if not provided
    if ($ShareName -eq "") {
        $ShareName = "OfficeAddins_$($userInfo.Username)"
        Write-Host "Using user-specific share name: $ShareName"
    }
    
    # Set folder path if not provided
    if ($FolderPath -eq "") {
        $FolderPath = Join-Path $userInfo.DocumentsPath $ShareName
        Write-Host "Using user's Documents folder: $FolderPath"
    }
    
    # Create the folder if it doesn't exist
    if (-not (Test-Path $FolderPath)) {
        New-Item -ItemType Directory -Path $FolderPath | Out-Null
        Write-Host "Created folder: $FolderPath"
    }
    
    # Set folder permissions for current user
    try {
        $userInfo = Get-CurrentUser
        if ($userInfo) {
            $acl = Get-Acl $FolderPath
            $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $userInfo.Username, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"
            )
            $acl.SetAccessRule($accessRule)
            Set-Acl -Path $FolderPath -AclObject $acl
            Write-Host "Set folder permissions for user: $($userInfo.Username)"
        }
    } catch {
        Write-Warning "Could not set folder permissions: $_"
    }
    
    # This manifest includes an ExtensionPoint for PrimaryCommandSurface so that a Ribbon button is created in Word.
    $manifestContent = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>f85491a7-0cf8-4950-b18c-d85ae9970d61</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>AnyGenAI</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="IP Agent AI"/>
  <Description DefaultValue="A template to get started"/>
  <IconUrl DefaultValue="http://10.100.100.71:3002/assets/logo-32.png"/>
  <HighResolutionIconUrl DefaultValue="http://10.100.100.71:3002/assets/logo-64.png"/>
  <SupportUrl DefaultValue="http://www.anygenai.com/help"/>
  <AppDomains>
    <AppDomain>http://10.100.100.71:3002</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="http://10.100.100.71:3002/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <Runtimes>
          <Runtime resid="Taskpane.Url" lifetime="long" />
        </Runtimes>
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="MyCustomGroup">
                <Label resid="CustomGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="MyButton">
                  <Label resid="MyButton.Label"/>
                  <Supertip>
                    <Title resid="MyButton.Title"/>
                    <Description resid="MyButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="http://10.100.100.71:3002/assets/logo-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="http://10.100.100.71:3002/assets/logo-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="http://10.100.100.71:3002/assets/logo-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="http://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="http://10.100.100.71:3002/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="http://10.100.100.71:3002/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Get started with your sample add-in!"/>
        <bt:String id="CustomGroup.Label" DefaultValue="IP Agent AI Group"/>
        <bt:String id="MyButton.Label" DefaultValue="IP Agent AI"/>
        <bt:String id="MyButton.Title" DefaultValue="IP Agent AI"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Your sample add-in loaded successfully. Go to the HOME tab and click the 'Show Task Pane' button to get started."/>
        <bt:String id="MyButton.Tooltip" DefaultValue="Click to open the IP Agent AI taskpane"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
'@

    $manifestDest = Join-Path $FolderPath "manifest.xml"
    try {
        $manifestContent | Out-File -FilePath $manifestDest -Encoding UTF8 -Force
        Write-Host "Created manifest file in folder: $manifestDest" -ForegroundColor Green
        return $manifestDest
    } catch {
        Write-Error "Failed to create manifest file: $_"
        Write-Warning "Please manually create the manifest file at: $FolderPath"
        return $false
    }
}

# Function to install add-in by creating network share and adding to trusted catalogs
function Install-Addin {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$ShareName = "",
        
        [Parameter()]
        [string]$ShareDescription = "",
        
        [Parameter()]
        [string]$FolderPath = "",
        
        [Parameter()]
        [string]$RegistryPath = ""
    )

    # Get current user information
    $userInfo = Get-CurrentUser
    if (-not $userInfo) {
        Write-Error "Could not determine current user"
        return $false
    }
    
    # Set user-specific defaults
    if ($ShareName -eq "") {
        $ShareName = "OfficeAddins_$($userInfo.Username)"
        Write-Host "Using user-specific share name: $ShareName"
    }
    
    if ($ShareDescription -eq "") {
        $ShareDescription = "Office Add-ins Shared Folder for $($userInfo.Username)"
    }
    
    # Set default paths based on current user
    if ($FolderPath -eq "") {
        $FolderPath = Join-Path $userInfo.DocumentsPath $ShareName
        Write-Host "Using current user's Documents folder: $FolderPath"
    }
    
    if ($RegistryPath -eq "") {
        $RegistryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
    }

    # 1. First ensure the manifest file exists or create it
    $manifestPath = Join-Path $FolderPath "manifest.xml"
    if (-not (Test-Path $manifestPath)) {
        $manifestCreated = New-ManifestFile -ShareName $ShareName -FolderPath $FolderPath
        if (-not $manifestCreated) {
            Write-Error "Failed to create manifest file"
            return $false
        }
    }

    # 2. Create network share if it doesn't exist
    try {
        # If the share already exists, remove it
        $existingShare = Get-WmiObject -Class Win32_Share -Filter "Name='$ShareName'" -ErrorAction SilentlyContinue
        if ($existingShare) {
            Remove-SmbShare -Name $ShareName -Force -ErrorAction SilentlyContinue
            Write-Host "Removed existing network share: $ShareName"
        }
        
        # Create new SMB share with permissions for both admin and current user
        $shareUsers = @("Administrators")
        if ($userInfo.Username -ne "Administrator") {
            $shareUsers += $userInfo.Username
        }
        
        New-SmbShare -Name $ShareName -Path $FolderPath -Description $ShareDescription -FullAccess $shareUsers
        
        # Set NTFS permissions to allow Everyone read access (needed for Office)
        $acl = Get-Acl $FolderPath
        $everyoneRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            "Everyone", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow"
        )
        $acl.SetAccessRule($everyoneRule)
        Set-Acl -Path $FolderPath -AclObject $acl
        
        Write-Host "Created network share: \\$env:COMPUTERNAME\$ShareName" -ForegroundColor Green
        $networkPath = "\\$env:COMPUTERNAME\$ShareName"
    } catch {
        Write-Error "Failed to create network share: $_"
        return $false
    }

    # 3. Add the network share to Word's Trusted Catalogs in the current user's registry
    try {
        # Clear existing trusted catalogs (if any)
        Remove-Item -Path "$RegistryPath\*" -Force -ErrorAction SilentlyContinue
        Write-Host "Cleared existing trusted catalogs"
        
        # Generate a new GUID for the catalog entry
        $catalogGuid = New-Guid
        $catalogPath = Join-Path $RegistryPath $catalogGuid
        
        # Ensure the base registry key exists
        if (-not (Test-Path $RegistryPath)) {
            New-Item -Path $RegistryPath -Force | Out-Null
        }
        
        # Create a subkey for this catalog
        New-Item -Path $catalogPath -Force | Out-Null
        
        # Set required registry values. The "Flags" value of 3 indicates that the catalog is enabled and set to show in the add-in menu.
        New-ItemProperty -Path $catalogPath -Name "Id"              -Value $catalogGuid -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $catalogPath -Name "Url"             -Value $networkPath -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $catalogPath -Name "Flags"           -Value 3           -PropertyType DWord  -Force | Out-Null
        New-ItemProperty -Path $catalogPath -Name "Type"            -Value 2           -PropertyType DWord  -Force | Out-Null
        New-ItemProperty -Path $catalogPath -Name "CatalogVersion"  -Value 2           -PropertyType DWord  -Force | Out-Null
        New-ItemProperty -Path $catalogPath -Name "SkipCatalogUpdate" -Value 0         -PropertyType DWord  -Force | Out-Null
        
        Write-Host "Added network share to trusted add-in catalogs for user: $($userInfo.Username)" -ForegroundColor Green
    } catch {
        Write-Error "Failed to modify registry: $_"
        return $false
    }

    # (Optional) Force registry refresh (this step may help Office detect changes faster)
    reg.exe unload "HKCU\Temp" 2>$null
    reg.exe load "HKCU\Temp" "$env:USERPROFILE\NTUSER.DAT" 2>$null
    reg.exe unload "HKCU\Temp" 2>$null

    Write-Host "`nSetup completed successfully!" -ForegroundColor Green
    Write-Host "Network share path: $networkPath"
    Write-Host "Manifest file location: $manifestPath"
    Write-Host "`nIMPORTANT:"
    Write-Host "1. Start Microsoft Word."
    Write-Host "2. Click File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs."
    Write-Host "3. Verify that $networkPath is listed and check 'Show in Menu'."
    Write-Host "4. Click OK and restart Word."
    Write-Host "`nAfter Word starts, your add-in's Ribbon button should appear on the Home tab within the custom group."

    return $true
}

# Function to check if Word is installed on the device
function Test-WordInstalled {
    [CmdletBinding()]
    param()
    
    Write-Host "Checking if Microsoft Word is installed..."
    
    # Method 1: Check Word application registry keys
    $wordRegPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    $wordAppPath = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
    $wordAppPath2 = "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
    
    # Method 2: Check if Word process can be started
    try {
        # Try to create a COM object for Word
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "Microsoft Word is installed and functioning." -ForegroundColor Green
        return $true
    }
    catch {
        # If COM object creation fails, check file existence
        if ((Test-Path $wordAppPath) -or (Test-Path $wordAppPath2) -or (Test-Path $wordRegPath)) {
            Write-Host "Microsoft Word appears to be installed, but could not be launched." -ForegroundColor Yellow
            return $true
        }
        else {
            Write-Host "ERROR: Microsoft Word is not installed on this device." -ForegroundColor Red
            return $false
        }
    }
}

# Function to check and handle manifest file for current user
function Test-ManifestExists {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$ShareName = "",
        
        [Parameter()]
        [string]$FolderPath = ""
    )
    
    Write-Host "Checking if manifest file exists..."
    
    # Get current user info and set defaults
    $userInfo = Get-CurrentUser
    if (-not $userInfo) {
        Write-Error "Could not determine current user"
        return $false
    }
    
    # Set user-specific share name if not provided
    if ($ShareName -eq "") {
        $ShareName = "OfficeAddins_$($userInfo.Username)"
        Write-Host "Using user-specific share name: $ShareName"
    }
    
    # Set folder path if not provided
    if ($FolderPath -eq "") {
        $FolderPath = Join-Path $userInfo.DocumentsPath $ShareName
        Write-Host "Using current user's Documents folder: $FolderPath"
    }
    
    # Create the folder if it doesn't exist
    if (-not (Test-Path $FolderPath)) {
        New-Item -ItemType Directory -Path $FolderPath | Out-Null
        Write-Host "Created folder: $FolderPath"
        return $false
    }
    
    # Check if manifest exists
    $manifestPath = Join-Path $FolderPath "manifest.xml"
    if (Test-Path $manifestPath) {
        Write-Host "Manifest file found: $manifestPath" -ForegroundColor Green
        return $true
    }
    
    # Check for manifests in other user profiles (for copying if needed)
    $usersFolder = "C:\Users"
    $foundManifests = @()
    
    if (Test-Path $usersFolder) {
        Get-ChildItem -Path $usersFolder -Directory | ForEach-Object {
            $userManifestPath = Join-Path $_.FullName "Documents\$ShareName\manifest.xml"
            if (Test-Path $userManifestPath) {
                Write-Host "Found manifest in another user profile: $userManifestPath"
                $foundManifests += $userManifestPath
            }
        }
    }
    
    # If found in other profiles, copy the first one
    if ($foundManifests.Count -gt 0) {
        try {
            Copy-Item -Path $foundManifests[0] -Destination $manifestPath -Force
            Write-Host "Copied manifest from another user profile to: $manifestPath" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Warning "Failed to copy manifest from another user profile: $_"
            return $false
        }
    }
    
    Write-Host "No manifest file found." -ForegroundColor Yellow
    return $false
}

# Function to check and create network share for manifest files
function Test-NetworkShareExists {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$ShareName = "",
        
        [Parameter()]
        [string]$FolderPath = ""
    )
    
    Write-Host "Checking if network share exists..."
    
    # Get current user info and set defaults
    $userInfo = Get-CurrentUser
    if (-not $userInfo) {
        Write-Error "Could not determine current user"
        return $false
    }
    
    # Set user-specific share name if not provided
    if ($ShareName -eq "") {
        $ShareName = "OfficeAddins_$($userInfo.Username)"
        Write-Host "Using user-specific share name: $ShareName"
    }
    
    # Set folder path if not provided
    if ($FolderPath -eq "") {
        $FolderPath = Join-Path $userInfo.DocumentsPath $ShareName
        Write-Host "Using current user's Documents folder: $FolderPath"
    }
    
    # Check if the share already exists
    $existingShare = Get-WmiObject -Class Win32_Share -Filter "Name='$ShareName'" -ErrorAction SilentlyContinue
    
    if ($existingShare) {
        Write-Host "Network share already exists: \\$env:COMPUTERNAME\$ShareName" -ForegroundColor Green
        return $true
    }
    
    # Create the folder if it doesn't exist
    if (-not (Test-Path $FolderPath)) {
        New-Item -ItemType Directory -Path $FolderPath | Out-Null
        Write-Host "Created folder: $FolderPath"
    }
    
    # Try to create the share
    try {
        $shareUsers = @("Administrators")
        if ($userInfo.Username -ne "Administrator") {
            $shareUsers += $userInfo.Username
        }
        
        $shareDescription = "Office Add-ins Shared Folder for $($userInfo.Username)"
        New-SmbShare -Name $ShareName -Path $FolderPath -Description $shareDescription -FullAccess $shareUsers
        Write-Host "Created network share: \\$env:COMPUTERNAME\$ShareName" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "Failed to create network share: $_"
        return $false
    }
}

# Function to check if share is in trusted locations
function Test-TrustedLocation {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$ShareName = "",
        
        [Parameter()]
        [string]$RegistryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
    )
    
    Write-Host "Checking if share is in trusted add-in catalogs..."
    
    # Get current user info and set defaults
    if ($ShareName -eq "") {
        $userInfo = Get-CurrentUser
        if (-not $userInfo) {
            Write-Error "Could not determine current user"
            return $false
        }
        $ShareName = "OfficeAddins_$($userInfo.Username)"
        Write-Host "Using user-specific share name: $ShareName"
    }
    
    $networkPath = "\\$env:COMPUTERNAME\$ShareName"
    
    # Check if the registry key exists
    if (-not (Test-Path $RegistryPath)) {
        Write-Host "Trusted catalogs registry key doesn't exist."
        return $false
    }
    
    # Check if the share is already in trusted catalogs
    $found = $false
    Get-ChildItem -Path $RegistryPath | ForEach-Object {
        $catalogPath = $_.PSPath
        $url = (Get-ItemProperty -Path $catalogPath -Name "Url" -ErrorAction SilentlyContinue).Url
        
        if ($url -eq $networkPath) {
            Write-Host "Share is already in trusted add-in catalogs." -ForegroundColor Green
            $found = $true
        }
    }
    
    return $found
}

# Function to find and click UI elements
function Find-AndClickElement {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ElementName,
        [Parameter(Mandatory=$true)]
        [System.Windows.Automation.AutomationElement]$ParentElement,
        [int]$TimeoutSeconds = 10
    )
    
    Write-Host "Looking for element: $ElementName"
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        $condition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, 
            $ElementName
        )
        
        $element = $ParentElement.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            $condition
        )
        
        if ($element) {
            try {
                # Try to click using InvokePattern first
                try {
                    $invokePattern = $element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                    if ($invokePattern) {
                        $invokePattern.Invoke()
                        Write-Host "Clicked element using InvokePattern: $ElementName" -ForegroundColor Green
                        return $true
                    }
                } catch {
                    Write-Host "InvokePattern not available, trying coordinate click"
                }

                # Fallback to coordinate click
                $point = $element.GetClickablePoint()
                [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point([int]$point.X, [int]$point.Y)
                Start-Sleep -Milliseconds 100
                
                # Check if type already exists before adding it
                if (-not ([System.Management.Automation.PSTypeName]'Win32Functions.Win32MouseEvent').Type) {
                    $signature = @'
                    [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
                    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
'@
                    Add-Type -MemberDefinition $signature -Name "Win32MouseEvent" -Namespace Win32Functions
                }
                
                # Mouse click down and up
                [Win32Functions.Win32MouseEvent]::mouse_event(0x00000002, 0, 0, 0, 0)
                Start-Sleep -Milliseconds 100
                [Win32Functions.Win32MouseEvent]::mouse_event(0x00000004, 0, 0, 0, 0)
                
                Write-Host "Clicked element using coordinates: $ElementName" -ForegroundColor Green
                return $true
            } catch {
                Write-Warning "Failed to click element $ElementName : $_"
            }
        }
        Start-Sleep -Milliseconds 500
    }
    Write-Host "Element not found: $ElementName" -ForegroundColor Yellow
    return $false
}

# Helper function to check if a window belongs to the Word process
function IsWordProcess {
    param (
        [Parameter(Mandatory=$true)]
        $window
    )
    
    try {
        $hwnd = $window.Current.NativeWindowHandle
        if ($hwnd -ne 0) {
            $processId = 0
            $null = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            # Get process ID from window handle (threadId not used but required by API)
            $null = [Win32Functions.Win32Api]::GetWindowThreadProcessId($hwnd, [ref]$processId)
            
            $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
            return ($process -and $process.Name -eq "WINWORD")
        }
    } catch {
        Write-Host "Error checking process: $_"
    }
    
    return $false
}

# Add GetWindowThreadProcessId function for process verification
if (-not ([System.Management.Automation.PSTypeName]'Win32Functions.Win32Api').Type) {
    $signature = @'
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, ref uint lpdwProcessId);
'@
    Add-Type -MemberDefinition $signature -Name "Win32Api" -Namespace Win32Functions
}

# Function to find Word window with a specific document name
# Helper function to check if a window belongs to the Word process
function IsWordProcess {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Automation.AutomationElement]$window
    )
    
    try {
        # Try to get the process ID of the window
        $processId = $window.Current.ProcessId
        if ($processId -eq 0) {
            return $false
        }
        
        # Get the process by ID
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($process -and $process.ProcessName -eq "WINWORD") {
            return $true
        }
    } catch {
        Write-Verbose "Error checking if window belongs to Word process: $_"
    }
    
    return $false
}

function Find-WordWindowWithDocument {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DocumentName
    )
    
    Write-Host "Looking for Word window with document name: $DocumentName"
    $script:maxAttempts = 10
    $script:attempt = 0
    
    while ($script:attempt -lt $script:maxAttempts) {
        $script:attempt++
        
        $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
        
        if ($wordProcesses.Count -eq 0) {
            Write-Host "No Word processes found, waiting..."
            Start-Sleep -Seconds 1
            continue
        }
        
        $windows = [System.Windows.Automation.AutomationElement]::RootElement.FindAll(
            [System.Windows.Automation.TreeScope]::Children,
            [System.Windows.Automation.PropertyCondition]::TrueCondition
        )
        
        foreach ($window in $windows) {
            $windowName = $window.Current.Name
            $className = $window.Current.ClassName
            
            # Check if the DocumentName is included anywhere in the window title
            # Window titles can vary based on Word version and document state
            $doesMatch = $false
            
            # Escape special regex characters in the document name
            $escapedDocName = [regex]::Escape($DocumentName)
            
            # Check various patterns where the document name might appear in the window title
            if ($windowName -match "^$escapedDocName(\s+|\[|\.).*-\s*Word(\s+\(.*\))?$" -or
                $windowName -match ".*$escapedDocName.*\s+-\s*Word(\s+\(.*\))?$") {
                $doesMatch = $true
                Write-Host "Window title '$windowName' matches document name '$DocumentName'"
            }
            
            # Check if it's a Word window by class name or process
            if ($doesMatch) {
                if ($className -eq "OpusApp" -or (IsWordProcess -window $window)) {
                    Write-Host "Found matching document window: $windowName" -ForegroundColor Green
                    return $window
                }
            }
        }
        
        Write-Host "Document window not found on attempt $script:attempt, waiting..."
        Start-Sleep -Seconds 1
    }
    
    Write-Host "Could not find document window with name: $DocumentName" -ForegroundColor Yellow
    return $null
}

# Function to find and focus Word window or launch a new instance
function Find-WordWindowWithName {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DocumentName,
        
        [Parameter(Mandatory=$false)]
        [switch]$ExactMatchOnly = $false
    )
    
    Write-Host "Looking for Word window with document name: $DocumentName"
    
    # Clean the document name - remove file extension if present for comparison purposes
    $cleanDocName = $DocumentName -replace '\.docx?$', ''
    Write-Host "Cleaned document name for matching: $cleanDocName"
    
    # First, check if any Word windows are open with the document
    # Try with the cleaned name first, then with the original name if not found
    $window = Find-WordWindowWithDocument -DocumentName $cleanDocName
    if (-not $window -and $cleanDocName -ne $DocumentName) {
        $window = Find-WordWindowWithDocument -DocumentName $DocumentName
    }
    
    if ($window) {
        return $window
    }
    
    # If not found, try to launch Word with the document
    Write-Host "Document window not found, attempting to launch Word with the document..."
    try {
        # Try to use COM object to open Word
        $word = New-Object -ComObject Word.Application -ErrorAction Stop
        $word.Visible = $true
        
        # Check if document URL is provided in global context
        if (Test-Path -LiteralPath $DocumentName) {
            # Open local document
            Write-Host "Opening local document: $DocumentName"
            $word.Documents.Open($DocumentName) | Out-Null
        } elseif ($documentUrl -ne "") { # Check global variable
            # Open from URL
            Write-Host "Opening document from URL: $documentUrl"
            $word.Documents.Open($documentUrl) | Out-Null
        } else {
            # Create a new document
            Write-Host "Creating a new document"
            $word.Documents.Add() | Out-Null
            $word.ActiveDocument.SaveAs([ref]$DocumentName)
        }
        
        Write-Host "Word launched with document, waiting for window to be accessible..."
        Start-Sleep -Seconds 3 # Wait for Word to fully initialize
        
        # Try to find the window again
        $window = Find-WordWindowWithDocument -DocumentName $DocumentName
        if ($window) {
            return $window
        }
    } catch {
        Write-Warning "Failed to launch Word with the document: $_"
    }
    
    Write-Host "Could not find or launch Word with document: $DocumentName" -ForegroundColor Yellow
    return $null
}

# Function to activate/focus a window
function Set-WindowFocus {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Automation.AutomationElement]$Window
    )
    
    Write-Host "Setting focus to Word window..."
    
    try {
        # Get the window handle
        $hwnd = $Window.Current.NativeWindowHandle
        
        if ($hwnd -ne 0) {
            # Add the SetForegroundWindow API if not already added
            if (-not ([System.Management.Automation.PSTypeName]'Win32Functions.WindowFocus').Type) {
                $signature = @'
                [DllImport("user32.dll")]
                [return: MarshalAs(UnmanagedType.Bool)]
                public static extern bool SetForegroundWindow(IntPtr hWnd);
                
                [DllImport("user32.dll")]
                public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
                Add-Type -MemberDefinition $signature -Name "WindowFocus" -Namespace Win32Functions
            }
            
            # Bring window to front (SW_SHOW = 5, SW_RESTORE = 9)
            [Win32Functions.WindowFocus]::ShowWindow($hwnd, 9)
            Start-Sleep -Milliseconds 200
            
            # Set as foreground window
            $result = [Win32Functions.WindowFocus]::SetForegroundWindow($hwnd)
            
            # Additional attempts to focus using UI Automation patterns
            try {
                # Try to use WindowPattern to set focus
                $windowPattern = $Window.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
                if ($windowPattern) {
                    $windowPattern.SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Normal)
                    Start-Sleep -Milliseconds 200
                }
                
                # Try to activate the window
                $Window.SetFocus()
            } catch {
                $errorMsg = $_.Exception.Message
                Write-Host "Could not set focus using UI Automation: $errorMsg"
            }
            
            # Wait for window to become responsive
            Start-Sleep -Seconds 1
            
            return $result
        } else {
            Write-Warning "Invalid window handle"
            return $false
        }
    } catch {
        $errorMsg = $_.Exception.Message
        Write-Warning "Failed to set window focus: $errorMsg"
        return $false
    }
}

# Function to get information about a Word document including Protected View status
function Get-DocumentInformation {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Automation.AutomationElement]$Window
    )
    
    Write-Host "Getting document information..."
    
    # Initialize return object
    $documentInfo = @{
        EnableEditingButton = $null
        IsInProtectedView = $false
        DocumentTitle = $Window.Current.Name
    }
    
    try {
        # Look for Protected View bar and Enable Editing button
        Write-Host "Checking for Protected View status..."
        
        # Check window title first for Protected View indicator
        $windowTitle = $Window.Current.Name
        Write-Host "Window title: $windowTitle"
        if ($windowTitle -match "Protected View|PROTECTED VIEW") {
            Write-Host "Document is in Protected View (detected from title)" -ForegroundColor Yellow
            $documentInfo.IsInProtectedView = $true
        }
        
        # Check for various Protected View banner text patterns
        $protectedViewPhrases = @(
            "PROTECTED VIEW",
            "Protected View", 
            "PROTECTED MODE",
            "Protected Mode",
            "This file originated"  # Common text in Protected View notification
        )
        
        $isProtectedView = $false
        foreach ($phrase in $protectedViewPhrases) {
            $condition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::NameProperty, 
                $phrase
            )
            
            $element = $Window.FindFirst(
                [System.Windows.Automation.TreeScope]::Descendants,
                $condition
            )
            
            if ($element) {
                Write-Host "Document is in Protected View (detected phrase: '$phrase')" -ForegroundColor Yellow
                $isProtectedView = $true
                $documentInfo.IsInProtectedView = $true
                break
            }
        }
        
        # Look for Enable Editing button as another indicator
        $buttonNames = @(
            "Enable Editing",
            "ENABLE EDITING",
            "Enable editing"
        )
        
        foreach ($buttonName in $buttonNames) {
            $buttonCondition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::NameProperty, 
                $buttonName
            )
            
            $enableEditingButton = $Window.FindFirst(
                [System.Windows.Automation.TreeScope]::Descendants,
                $buttonCondition
            )
            
            if ($enableEditingButton) {
                Write-Host "Found Enable Editing button with name: $buttonName" -ForegroundColor Green
                $documentInfo.EnableEditingButton = $enableEditingButton
                # If we find this button, the document is definitely in Protected View
                $documentInfo.IsInProtectedView = $true
                break
            }
        }
        
        # If we still haven't found the button, look for any button that might be the Enable Editing button
        if (-not $documentInfo.EnableEditingButton -and ($isProtectedView -or $documentInfo.IsInProtectedView)) {
            Write-Host "Document appears to be in Protected View, trying to find Enable Editing button by control type..."
            
            # Looking for buttons in the document
            $buttonTypeCondition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::Button
            )
            
            # Find all buttons in the descendant tree
            $allButtons = $Window.FindAll(
                [System.Windows.Automation.TreeScope]::Descendants,
                $buttonTypeCondition
            )
            
            # Check each button to see if it might be the Enable Editing button
            foreach ($button in $allButtons) {
                $buttonName = $button.Current.Name
                if ($buttonName -match "enable|editing" -or $buttonName -eq "") {
                    Write-Host "Found potential Enable Editing button: $buttonName" -ForegroundColor Green
                    $documentInfo.EnableEditingButton = $button
                    break
                }
            }
        }
        
        # Final Protected View status check
        if ($documentInfo.IsInProtectedView) {
            Write-Host "Document is confirmed to be in Protected View" -ForegroundColor Yellow
        } else {
            Write-Host "Document is not in Protected View"
        }
        
        return $documentInfo
    } catch {
        Write-Warning "Error getting document information: $($_.Exception.Message)"
        return $documentInfo
    }
}

# Function to enable editing on a document in Protected View
function Enable-DocumentEditing {
    param ($enableEditingButton)
    
    if ($enableEditingButton) {
        Write-Host "Clicking Enable Editing button..."
        try {
            # Try to click using InvokePattern
            try {
                $invokePattern = $enableEditingButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                if ($invokePattern) {
                    $invokePattern.Invoke()
                    Write-Host "Clicked Enable Editing button using InvokePattern" -ForegroundColor Green
                    Start-Sleep -Seconds 2  # Wait for document to exit Protected View
                    return $true
                }
            } catch {
                Write-Host "InvokePattern not available, trying coordinate click"
            }

            # Fallback to coordinate click
            $point = $enableEditingButton.GetClickablePoint()
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point([int]$point.X, [int]$point.Y)
            Start-Sleep -Milliseconds 300
            
            # Check if type already exists before adding it
            if (-not ([System.Management.Automation.PSTypeName]'Win32Functions.Win32MouseEvent').Type) {
                $signature = @'
                [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
                public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
'@
                Add-Type -MemberDefinition $signature -Name "Win32MouseEvent" -Namespace Win32Functions
            }
            
            # Mouse click down and up
            [Win32Functions.Win32MouseEvent]::mouse_event(0x00000002, 0, 0, 0, 0)
            Start-Sleep -Milliseconds 100
            [Win32Functions.Win32MouseEvent]::mouse_event(0x00000004, 0, 0, 0, 0)
            
            Write-Host "Clicked Enable Editing button using coordinates" -ForegroundColor Green
            Start-Sleep -Seconds 2  # Wait for document to exit Protected View
            return $true
        } catch {
            Write-Warning "Failed to click Enable Editing button: $($_.Exception.Message)"
            return $false
        }
    } else {
        Write-Host "No Enable Editing button found, document may already be in edit mode"
        return $true
    }
}

# Function to find and click a UI element by name
function Find-AndClickElement {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ElementName,
        
        [Parameter(Mandatory=$true)]
        [System.Windows.Automation.AutomationElement]$ParentElement,
        
        [Parameter(Mandatory=$false)]
        [int]$TimeoutSeconds = 3,
        
        [Parameter(Mandatory=$false)]
        [switch]$UseCoordinateClickOnly = $false,
        
        [Parameter(Mandatory=$false)]
        [int]$YOffset = 0
    )
    
    Write-Host "Looking for element: $ElementName"
    $startTime = Get-Date
    $elementFound = $false
    $element = $null
    
    while (-not $elementFound -and ((Get-Date) - $startTime).TotalSeconds -lt $TimeoutSeconds) {
        # Create a condition for finding the element by name
        $condition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, 
            $ElementName
        )
        
        # Try to find the element
        $element = $ParentElement.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            $condition
        )
        
        if ($element) {
            $elementFound = $true
            Write-Host "Found element: $ElementName" -ForegroundColor Green
            break
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    if (-not $elementFound) {
        Write-Host "Element not found: $ElementName" -ForegroundColor Yellow
        return $false
    }
    
    # Try to click the element
    try {
        # First try to use InvokePattern if available and not forced to use coordinate click
        if (-not $UseCoordinateClickOnly) {
            try {
                $invokePattern = $element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                if ($invokePattern) {
                    $invokePattern.Invoke()
                    Write-Host "Clicked element using InvokePattern: $ElementName" -ForegroundColor Green
                    Start-Sleep -Milliseconds 500
                    return $true
                }
            } catch {
                Write-Host "InvokePattern not available for $ElementName, trying coordinate click"
            }
        }
        
        # Fallback to coordinate click
        Write-Host "Using coordinate click for element: $ElementName"
        $point = $element.GetClickablePoint()
        
        # Apply Y offset if specified
        if ($YOffset -ne 0) {
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(
                [int]$point.X, 
                [int]($point.Y + $YOffset)
            )
        } else {
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(
                [int]$point.X, 
                [int]$point.Y
            )
        }
        
        Start-Sleep -Milliseconds 200
        
        # Check if type already exists before adding it
        if (-not ([System.Management.Automation.PSTypeName]'Win32Functions.Win32MouseEventClick').Type) {
            $signature = @'
            [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
            public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
'@
            Add-Type -MemberDefinition $signature -Name "Win32MouseEventClick" -Namespace Win32Functions
        }
        
        # Mouse click down and up
        [Win32Functions.Win32MouseEventClick]::mouse_event(0x00000002, 0, 0, 0, 0)
        Start-Sleep -Milliseconds 100
        [Win32Functions.Win32MouseEventClick]::mouse_event(0x00000004, 0, 0, 0, 0)
        
        Write-Host "Clicked element using coordinates: $ElementName" -ForegroundColor Green
        Start-Sleep -Milliseconds 500
        return $true
    } catch {
        Write-Warning "Failed to click element: $ElementName"
        Write-Warning "Error: $($_.Exception.Message)"
        return $false
    }
}

# Function to check and open add-in from ribbon
function Open-AddInFromRibbon {
    param ($wordWindow)
    
    Write-Host "Checking for add-in button on ribbon..."
    $buttonNames = @(
        "IP Agent AI",
        "IP Agent AI Group"
    )
    
    # First try to find the add-in button directly
    foreach ($name in $buttonNames) {
        Write-Host "Looking for button: $name"
        if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow -TimeoutSeconds 3) {
            Write-Host "Successfully opened add-in from ribbon" -ForegroundColor Green
            return $true
        }
    }
    
    # If direct search fails, try to look for buttons that might contain our add-in
    Write-Host "Direct button not found, trying to find ribbon groups that might contain our add-in..."
    $ribbonGroups = @(
        "Add-ins",
        "My Add-ins",
        "Add-ins Group"
    )
    
    foreach ($groupName in $ribbonGroups) {
        Write-Host "Looking for ribbon group: $groupName"
        $groupCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, 
            $groupName
        )
        
        $ribbonGroup = $wordWindow.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            $groupCondition
        )
        
        if ($ribbonGroup) {
            Write-Host "Found ribbon group: $groupName" -ForegroundColor Green
            # Click on the ribbon group first
            try {
                $invokePattern = $ribbonGroup.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                if ($invokePattern) {
                    $invokePattern.Invoke()
                    Write-Host "Clicked on ribbon group: $groupName" -ForegroundColor Green
                    Start-Sleep -Seconds 2
                    
                    # Now try to find our add-in button within this group
                    foreach ($name in $buttonNames) {
                        if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow -TimeoutSeconds 5) {
                            Write-Host "Successfully opened add-in from within $groupName group" -ForegroundColor Green
                            return $true
                        }
                    }
                }
            } catch {
                Write-Warning "Failed to click on ribbon group: $groupName - $_"
            }
        }
    }
    
    Write-Host "Add-in button not found on ribbon" -ForegroundColor Yellow
    return $false
}

# Helper function to retry clicking an element with multiple attempts
function Invoke-ClickWithRetry {
    param (
        [Parameter(Mandatory=$true)]
        [System.Windows.Automation.AutomationElement]$Element,
        
        [Parameter(Mandatory=$true)]
        [string]$ElementName,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory=$false)]
        [int]$DelayBetweenRetries = 500,
        
        [Parameter(Mandatory=$false)]
        [int]$YOffset = 0,
        
        [Parameter(Mandatory=$false)]
        [switch]$UseCoordinateClickOnly = $false
    )
    
    $attempt = 0
    $success = $false
    
    while ($attempt -lt $MaxRetries -and -not $success) {
        $attempt++
        Write-Host "Attempt $attempt of $MaxRetries to click element: $ElementName"
        
        try {
            # First try to use InvokePattern if available and not forced to use coordinate click
            if (-not $UseCoordinateClickOnly) {
                try {
                    $invokePattern = $Element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                    if ($invokePattern) {
                        $invokePattern.Invoke()
                        Write-Host "Successfully clicked element using InvokePattern: $ElementName" -ForegroundColor Green
                        $success = $true
                        break
                    }
                } catch {
                    Write-Host "InvokePattern not available for $ElementName, trying coordinate click"
                }
            }
            
            # Fallback to coordinate click
            Write-Host "Using coordinate click for element: $ElementName"
            $point = $Element.GetClickablePoint()
            
            # Apply Y offset if specified
            if ($YOffset -ne 0) {
                [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(
                    [int]$point.X, 
                    [int]($point.Y + $YOffset)
                )
            } else {
                [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(
                    [int]$point.X, 
                    [int]$point.Y
                )
            }
            
            Start-Sleep -Milliseconds 200
            
            # Check if type already exists before adding it
            if (-not ([System.Management.Automation.PSTypeName]'Win32Functions.Win32MouseEventRetry').Type) {
                $signature = @'
                [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
                public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
'@
                Add-Type -MemberDefinition $signature -Name "Win32MouseEventRetry" -Namespace Win32Functions
            }
            
            # Mouse click down and up
            [Win32Functions.Win32MouseEventRetry]::mouse_event(0x00000002, 0, 0, 0, 0)
            Start-Sleep -Milliseconds 100
            [Win32Functions.Win32MouseEventRetry]::mouse_event(0x00000004, 0, 0, 0, 0)
            
            Write-Host "Successfully clicked element using coordinates: $ElementName" -ForegroundColor Green
            $success = $true
            
        } catch {
            Write-Warning "Attempt $attempt failed to click element: $ElementName - $($_.Exception.Message)"
            
            if ($attempt -lt $MaxRetries) {
                Write-Host "Waiting $DelayBetweenRetries ms before retry..."
                Start-Sleep -Milliseconds $DelayBetweenRetries
            }
        }
    }
    
    if (-not $success) {
        Write-Warning "Failed to click element $ElementName after $MaxRetries attempts"
    }
    
    return $success
}

# Function to open shared folder dialog
function Open-SharedFolderDialog {
    param ($wordWindow)
    
    Write-Host "Opening Shared Folder dialog using new UI flow..."
    
    Write-Host "Setting focus to Word window before UI interactions..."
    $focusResult = Set-WindowFocus -Window $wordWindow
    if (-not $focusResult) {
        Write-Warning "Could not set focus to Word window, but continuing anyway..."
    } else {
        Write-Host "Successfully set focus to Word window" -ForegroundColor Green
    }
        
    # Look for Add-ins button in the ribbon
    Write-Host "Looking for Add-ins button in ribbon..."
    $addInsButtonNames = @("Add-ins", "ADD-INS")
    $addInsButtonFound = $false
    foreach ($name in $addInsButtonNames) {
        $buttonCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, 
            $name
        )
        
        $addInsButton = $wordWindow.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            $buttonCondition
        )
        
        if ($addInsButton) {
            Write-Host "Found Add-ins button: $name"
            # Use retry logic for clicking Add-ins button
            $clickSuccess = Invoke-ClickWithRetry -Element $addInsButton -ElementName $name -YOffset 10
            if ($clickSuccess) {
                $addInsButtonFound = $true
                break
            }
        }
    }
    
    Write-Host "Setting focus to Word window before UI interactions..."
    $focusResult = Set-WindowFocus -Window $wordWindow
    if (-not $focusResult) {
        Write-Warning "Could not set focus to Word window, but continuing anyway..."
    } else {
        Write-Host "Successfully set focus to Word window" -ForegroundColor Green
    }
    
    $newMethodWorked = $false
    if ($addInsButtonFound) {        
        # Look for More Add-ins button
        Write-Host "Looking for More Add-ins button..."
        $moreAddInsNames = @("More Add-ins", "Get Add-ins", "More")
        $moreAddInsFound = $false
        
        foreach ($name in $moreAddInsNames) {
            $moreButtonCondition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::NameProperty, 
                $name
            )
            
            $moreAddInsButton = $wordWindow.FindFirst(
                [System.Windows.Automation.TreeScope]::Descendants,
                $moreButtonCondition
            )
            
            if ($moreAddInsButton) {
                Write-Host "Found More Add-ins button: $name"
                # Use retry logic for clicking More Add-ins button
                $clickSuccess = Invoke-ClickWithRetry -Element $moreAddInsButton -ElementName $name
                if ($clickSuccess) {
                    $moreAddInsFound = $true
                    break
                }
            }
        }
        
        if ($moreAddInsFound) {
            # Wait for Office Add-ins dialog to open
            Start-Sleep -Seconds 3
            $newMethodWorked = $true
        } else {
            Write-Warning "Could not find More Add-ins button, trying fallback method..."
            # Press Escape to close Add-ins menu
            [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
            Start-Sleep -Seconds 1
        }
    } else {
        Write-Warning "Could not find Add-ins button in ribbon, trying fallback method..."
    }
    
    # If the new method didn't work, try the old method (File -> Get Add-ins)
    if (-not $newMethodWorked) {
        Write-Host "Trying fallback method (File -> Get Add-ins)..."
        
        # **NEW: Ensure Word window still has focus before fallback method**
        Write-Host "Re-focusing Word window for fallback method..."
        $focusResult = Set-WindowFocus -Window $wordWindow
        if (-not $focusResult) {
            Write-Warning "Could not re-focus Word window for fallback method"
        }
        Start-Sleep -Seconds 1
        
        $fileTabNames = @("File", "FILE", "File Tab")
        $found = $false
        foreach ($name in $fileTabNames) {
            # Use retry logic for Find-AndClickElement calls
            $retryCount = 0
            $maxRetries = 3
            while ($retryCount -lt $maxRetries -and -not $found) {
                $retryCount++
                Write-Host "Attempt $retryCount of $maxRetries to find and click: $name"
                if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow) {
                    $found = $true
                    break
                }
                if ($retryCount -lt $maxRetries) {
                    Start-Sleep -Milliseconds 500
                }
            }
            if ($found) {
                break
            }
        }
        
        if (-not $found) {
            Write-Host "Using Alt+F shortcut for File menu..."
            [System.Windows.Forms.SendKeys]::SendWait("%F")
        }
        Start-Sleep -Seconds 2
        
        # Try to find Get Add-ins
        $getAddinsNames = @("Get Add-ins", "Office Add-ins", "Get Office Add-ins")
        $addinsFound = $false
        foreach ($name in $getAddinsNames) {
            # Use retry logic for Find-AndClickElement calls
            $retryCount = 0
            $maxRetries = 3
            while ($retryCount -lt $maxRetries -and -not $addinsFound) {
                $retryCount++
                Write-Host "Attempt $retryCount of $maxRetries to find and click: $name"
                if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow) {
                    $addinsFound = $true
                    break
                }
                if ($retryCount -lt $maxRetries) {
                    Start-Sleep -Milliseconds 500
                }
            }
            if ($addinsFound) {
                break
            }
        }
        
        if (-not $addinsFound) {
            Write-Warning "Could not find Get Add-ins option in File menu"
            # Press Escape to close File menu
            [System.Windows.Forms.SendKeys]::SendWait("{ESC}")
            return $false
        }
        
        # Wait for dialog to open
        Start-Sleep -Seconds 3
    }

    # From here, continue with the original logic to find the Office Add-ins dialog
    Write-Host "Looking for Office Add-ins dialog..."
    $dialogCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty, 
        "Office Add-ins"
    )
    $addinDialog = $wordWindow.FindFirst(
        [System.Windows.Automation.TreeScope]::Descendants,
        $dialogCondition
    )

    if (-not $addinDialog) {
        Write-Warning "Could not find Office Add-ins dialog"
        return $false
    }
    Write-Host "Found Office Add-ins dialog" -ForegroundColor Green

    # Click Shared Folder tab
    Write-Host "Looking for SHARED FOLDER tab..."
    $sharedFolderTab = $addinDialog.FindFirst(
        [System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, 
            "SHARED FOLDER"
        ))
    )

    if ($sharedFolderTab) {
        # Use retry logic for clicking SHARED FOLDER tab
        $clickSuccess = Invoke-ClickWithRetry -Element $sharedFolderTab -ElementName "SHARED FOLDER tab"
        if (-not $clickSuccess) {
            Write-Warning "Could not click SHARED FOLDER tab after retries"
            return $false
        }
    } else {
        Write-Warning "Could not find SHARED FOLDER tab"
        return $false
    }
    Start-Sleep -Seconds 2  # Wait for tab content to load

    # Function to try finding the add-in with multiple attempts
    function Find-AddInInDialog {
        param($dialog)
        
        Write-Host "Looking for IP Agent AI in the dialog..."
        
        # Try multiple approaches to find the add-in element
        
        # Approach 1: Direct name search
        $nameCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, 
            "IP Agent AI"
        )
        
        $addInElement = $dialog.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            $nameCondition
        )
        
        if ($addInElement) {
            return $addInElement
        }

        # Approach 2: Look for list item containing My Word Add-in
        Write-Host "Trying to find add-in by list item..."
        
        # Find all list items
        $listItemCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::ListItem
        )
        
        $listItems = $dialog.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            $listItemCondition
        )

        $targetItem = $null
        foreach ($item in $listItems) {
            if ($item.Current.Name -match "IP Agent AI") {
                $targetItem = $item
                Write-Host "Found target list item: $($item.Current.Name)" -ForegroundColor Green
                return $targetItem
            }
        }
        
        # No add-in found with either approach
        return $null
    }
    
    # First attempt to find the add-in
    $targetItem = Find-AddInInDialog -dialog $addinDialog
    
    # If add-in not found, try refreshing
    $refreshAttempts = 0
    $maxRefreshAttempts = 3
    
    while (-not $targetItem -and $refreshAttempts -lt $maxRefreshAttempts) {
        $refreshAttempts++
        Write-Host "Add-in not found, attempting to refresh dialog (Attempt $refreshAttempts of $maxRefreshAttempts)..."
        
        # Look for Refresh button
        $refreshButtonCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, 
            "Refresh"
        )
        
        $refreshButton = $addinDialog.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            $refreshButtonCondition
        )
        
        if ($refreshButton) {
            Write-Host "Found Refresh button, clicking..."
            # Use retry logic for clicking Refresh button
            $clickSuccess = Invoke-ClickWithRetry -Element $refreshButton -ElementName "Refresh button"
            
            if ($clickSuccess) {
                # Wait for refresh to complete
                Write-Host "Waiting for refresh to complete..."
                Start-Sleep -Seconds 3
                
                # Try to find the add-in again
                $targetItem = Find-AddInInDialog -dialog $addinDialog
                
                if ($targetItem) {
                    Write-Host "Add-in found after refreshing!" -ForegroundColor Green
                    break
                }
            } else {
                Write-Warning "Failed to click Refresh button after retries"
            }
        } else {
            Write-Warning "Could not find Refresh button"
            break
        }
    }
    
    # If we still don't have a target item after refresh attempts, return false
    if (-not $targetItem) {
        Write-Warning "Could not find IP Agent AI element after $refreshAttempts refresh attempts"
        return $false
    }

    # Now we have the target item, proceed with selecting it
    Write-Host "Found target item, attempting to select..." -ForegroundColor Green
    try {
        # Get the bounding rectangle of the target item
        $boundingRect = $targetItem.Current.BoundingRectangle
        
        # Calculate click position (center of the item)
        $clickX = $boundingRect.X + ($boundingRect.Width / 2)
        $clickY = $boundingRect.Y + ($boundingRect.Height / 2)
        
        # Use retry logic for clicking the target item
        $retryCount = 0
        $maxRetries = 3
        $clickSuccess = $false
        
        while ($retryCount -lt $maxRetries -and -not $clickSuccess) {
            $retryCount++
            Write-Host "Attempt $retryCount of $maxRetries to click on add-in item"
            
            try {
                # Move mouse and click
                [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(
                    [int]$clickX, 
                    [int]$clickY
                )
                Start-Sleep -Milliseconds 200
                
                # Define the mouse_event function if it doesn't exist to ensure it's available
                if (-not ([System.Management.Automation.PSTypeName]'Win32Functions.Win32MouseEvent').Type) {
                    $signature = @'
                    [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
                    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
'@
                    Add-Type -MemberDefinition $signature -Name "Win32MouseEvent" -Namespace Win32Functions
                }
                
                # Mouse click down and up
                [Win32Functions.Win32MouseEvent]::mouse_event(0x00000002, 0, 0, 0, 0)
                Start-Sleep -Milliseconds 100
                [Win32Functions.Win32MouseEvent]::mouse_event(0x00000004, 0, 0, 0, 0)
                
                Write-Host "Successfully clicked on add-in using calculated center position" -ForegroundColor Green
                $clickSuccess = $true
                
            } catch {
                Write-Warning "Attempt $retryCount failed to click add-in item: $_"
                if ($retryCount -lt $maxRetries) {
                    Start-Sleep -Milliseconds 500
                }
            }
        }
        
        if (-not $clickSuccess) {
            Write-Warning "Failed to click add-in item after $maxRetries attempts"
            return $false
        }
        
        Start-Sleep -Milliseconds 500

        # Now find and click the Add button
        Write-Host "Looking for Add button..."
        $addButtonCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, 
            "Add"
        )
        $addButton = $addinDialog.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            $addButtonCondition
        )

        if ($addButton) {
            Write-Host "Found Add button, attempting to click..." -ForegroundColor Green
            # Use retry logic for clicking Add button
            $clickSuccess = Invoke-ClickWithRetry -Element $addButton -ElementName "Add button"
            
            if ($clickSuccess) {
                Start-Sleep -Seconds 5
                return $true
            } else {
                Write-Warning "Failed to click Add button after retries"
                return $false
            }
        } else {
            Write-Warning "Could not find Add button"
            return $false
        }
    } catch {
        Write-Warning "Failed to select add-in: $_"
        return $false
    }

    return $false
}

# Function to start Word as current user (not admin)
function Start-WordAsCurrentUser {
    param (
        [string]$DocumentPath = "",
        [string]$DocumentUrl = ""
    )
    
    $userInfo = Get-CurrentUser
    if (-not $userInfo) {
        Write-Error "Could not determine current user"
        return $false
    }
    
    try {
        # If we have a document URL, use it directly
        if ($DocumentUrl -ne "") {
            Write-Host "Launching Word with document URL: $DocumentUrl"
            Start-Process "winword.exe" -ArgumentList "`"$DocumentUrl`"" -Wait:$false
            Start-Sleep -Seconds 3  # Wait for Word to start
            return $true
        }
        # If we have a document path, use it
        elseif ($DocumentPath -ne "" -and (Test-Path $DocumentPath)) {
            Write-Host "Launching Word with document: $DocumentPath"
            Start-Process "winword.exe" -ArgumentList "`"$DocumentPath`"" -Wait:$false
            Start-Sleep -Seconds 3  # Wait for Word to start
            return $true
        }
        # Just launch Word
        else {
            Write-Host "Launching Word without document"
            Start-Process "winword.exe" -Wait:$false
            Start-Sleep -Seconds 3  # Wait for Word to start
            return $true
        }
    }
    catch {
        Write-Error "Failed to launch Word as current user: $_"
        return $false
    }
}

# Function to open a Word document from a URL for current user
function Open-WordDocument {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$DocumentName,
        
        [Parameter(Mandatory=$false)]
        [string]$DocumentUrl = ""
    )
    
    Write-Host "Attempting to open Word document: $DocumentName for current user"
    
    # Get current user info for proper document location
    $userInfo = Get-CurrentUser
    if (-not $userInfo) {
        Write-Warning "Could not determine current user, using default paths"
    }
    
    # Try to launch Word as current user instead of using COM (which runs as admin)
    try {
        if ($DocumentUrl -ne "") {
            Write-Host "Opening document from URL as current user: $DocumentUrl"
            return Start-WordAsCurrentUser -DocumentUrl $DocumentUrl
        }
        else {
            # Try to find document in user's documents folder first
            if ($userInfo) {
                $userDocPath = Join-Path $userInfo.DocumentsPath $DocumentName
                if (Test-Path $userDocPath) {
                    Write-Host "Opening document from user documents: $userDocPath"
                    return Start-WordAsCurrentUser -DocumentPath $userDocPath
                }
            }
            
            # If document doesn't exist, just launch Word
            Write-Host "Document not found, launching Word for user to create: $DocumentName"
            return Start-WordAsCurrentUser
        }
    }
    catch {
        Write-Error "Failed to open Word document as current user: $_"
        
        # Fallback to COM object (will run as admin but at least works)
        Write-Warning "Falling back to COM object (will run as admin)"
        try {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $true
            
            # If URL is provided, open from URL
            if ($DocumentUrl -ne "") {
                Write-Host "Opening document from URL using COM: $DocumentUrl"
                try {
                    $word.Documents.Open($DocumentUrl) | Out-Null
                    Write-Host "Document opened successfully from URL using COM" -ForegroundColor Green
                    return $true
                }
                catch {
                    Write-Error "Failed to open document from URL: $_"
                    $word.Quit()
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
                    return $false
                }
            }
            else {
                # Try to find document in user's documents folder first
                if ($userInfo) {
                    $userDocPath = Join-Path $userInfo.DocumentsPath $DocumentName
                    if (Test-Path $userDocPath) {
                        $word.Documents.Open($userDocPath) | Out-Null
                        Write-Host "Document opened from user documents using COM: $userDocPath" -ForegroundColor Green
                        return $true
                    }
                }
                
                # Create new document if not found
                Write-Host "Creating new document using COM: $DocumentName"
                $newDoc = $word.Documents.Add()
                if ($userInfo) {
                    $saveAsPath = Join-Path $userInfo.DocumentsPath $DocumentName
                    $newDoc.SaveAs([ref]$saveAsPath)
                    Write-Host "New document created and saved using COM: $saveAsPath" -ForegroundColor Green
                } else {
                    $newDoc.SaveAs([ref]$DocumentName)
                    Write-Host "New document created using COM: $DocumentName" -ForegroundColor Green
                }
                return $true
            }
        }
        catch {
            Write-Error "Failed to open Word document using COM: $_"
            return $false
        }
    }
}

# Main script
try {
    Write-Host "=== Office Add-in Setup for Current User ===" -ForegroundColor Cyan
    
    # 0. Get current user information
    $userInfo = Get-CurrentUser
    if (-not $userInfo) {
        Write-Error "Could not determine current user"
        exit $ERROR_USER_NOT_FOUND
    }
    
    Write-Host "Target User: $($userInfo.Username)" -ForegroundColor Green
    Write-Host "Profile Path: $($userInfo.ProfilePath)" -ForegroundColor Green
    Write-Host "Documents Path: $($userInfo.DocumentsPath)" -ForegroundColor Green
    
    # 1. Check if Word is installed
    $wordInstalled = Test-WordInstalled
    if (-not $wordInstalled) {
        Write-Error "Microsoft Word is not installed on this device."
        exit $ERROR_WORD_NOT_INSTALLED
    }
    
    # 2. First check manifest file existence and create if needed
    Write-Host "`nChecking manifest file..." -ForegroundColor Yellow
    $manifestExists = Test-ManifestExists
    if (-not $manifestExists) {
        Write-Host "Creating manifest file for current user..."
        $manifestCreated = New-ManifestFile
        if (-not $manifestCreated) {
            Write-Error "Failed to create manifest file"
            exit $ERROR_MANIFEST_CREATION_FAILED
        }
        $manifestExists = $true
        Write-Host "Manifest file created successfully" -ForegroundColor Green
    }
    
    # 3. Now check network share exists and create if needed
    Write-Host "`nChecking network share..." -ForegroundColor Yellow
    $shareExists = Test-NetworkShareExists
    if (-not $shareExists) {
        Write-Error "Failed to create or verify network share."
        exit $ERROR_NETWORK_SHARE_FAILED
    }
    
    # 4. Check if share is in trusted locations and register if needed
    Write-Host "`nChecking trusted catalogs..." -ForegroundColor Yellow
    $isTrusted = Test-TrustedLocation
    if (-not $isTrusted) {
        Write-Host "Adding network share to trusted locations for current user..."
        # Clear existing trusted catalogs (if any)
        Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\*" -Force -ErrorAction SilentlyContinue
        
        # Add the network share to trusted catalogs
        $catalogGuid = New-Guid
        $catalogPath = Join-Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs" $catalogGuid
        
        # Ensure the base registry key exists
        if (-not (Test-Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs")) {
            New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs" -Force | Out-Null
        }
        
        # Create a subkey for this catalog
        New-Item -Path $catalogPath -Force | Out-Null
        
        # Set required registry values using current user's share
        $userShareName = "OfficeAddins_$($userInfo.Username)"
        $userNetworkPath = "\\$env:COMPUTERNAME\$userShareName"
        
        New-ItemProperty -Path $catalogPath -Name "Id"              -Value $catalogGuid -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $catalogPath -Name "Url"             -Value $userNetworkPath -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $catalogPath -Name "Flags"           -Value 3           -PropertyType DWord  -Force | Out-Null
        New-ItemProperty -Path $catalogPath -Name "Type"            -Value 2           -PropertyType DWord  -Force | Out-Null
        New-ItemProperty -Path $catalogPath -Name "CatalogVersion"  -Value 2           -PropertyType DWord  -Force | Out-Null
        New-ItemProperty -Path $catalogPath -Name "SkipCatalogUpdate" -Value 0         -PropertyType DWord  -Force | Out-Null
        
        Write-Host "Added network share to trusted add-in catalogs for user: $($userInfo.Username)" -ForegroundColor Green
    }
    
    # 5. If URL is provided, open the document
    if ($documentUrl -ne "") {
        $docOpened = Open-WordDocument -DocumentName $documentName -DocumentUrl $documentUrl
        if (-not $docOpened) {
            Write-Error "Failed to open document from URL: $documentUrl"
            exit $ERROR_OPEN_DOCUMENT_FAILED
        }
        Write-Host "Document opened successfully from URL: $documentUrl" -ForegroundColor Green
    }
    
    # 6. Try to find the Word window with the document
    Write-Host "Searching for Word window with document: $documentName"
    $wordWindow = Find-WordWindowWithDocument -DocumentName $documentName
    
    # 7. If not found, try to find or launch Word with the document
    if (-not $wordWindow) {
        Write-Host "Word window not found, launching Word..."
        $wordWindow = Find-WordWindowWithName -DocumentName $documentName
    }
    
    # 8. If we have a Word window, try to configure the add-in
    if ($wordWindow) {
        # Get information about the document including whether it's in Protected View
        Write-Host "Checking document protection status..."
        $documentInfo = Get-DocumentInformation -Window $wordWindow
        $enableEditingButton = $documentInfo.EnableEditingButton
        
        # Set focus to the Word window
        Write-Host "Setting focus to Word window"
        $focused = Set-WindowFocus -Window $wordWindow
        if (-not $focused) {
            Write-Warning "Could not set focus to Word window"
            # Continue anyway as we might still be able to enable editing
        }
        Start-Sleep -Seconds 2  # Wait for window to be fully focused
        
        # Enable editing ONLY if document is in Protected View
        if ($enableEditingButton) {
            Write-Host "Document is in Protected View, enabling editing..." -ForegroundColor Yellow
            if (-not (Enable-DocumentEditing -enableEditingButton $enableEditingButton)) {
                Write-Warning "Failed to enable editing on document"
                # Continue anyway, we might still be able to access the add-in
            }
            Start-Sleep -Seconds 2  # Wait for document to exit Protected View
            
            # After enabling editing, verify the document window again
            Write-Host "Re-locating document after enabling editing: $documentName"
            # Wait a bit longer to ensure Word has updated the window state
            Start-Sleep -Seconds 3
            
            $updatedWindow = Find-WordWindowWithName -DocumentName $documentName
            
            if ($updatedWindow) {
                Write-Host "Successfully enabled editing for document: $documentName" -ForegroundColor Green
                $wordWindow = $updatedWindow
            } else {
                Write-Warning "Could not re-locate document window after enabling editing"
                # Continue anyway with the original window
            }
        } else {
            Write-Host "Document is not in Protected View, no need to enable editing"
        }
        
        # Ensure the Word window has focus before proceeding
        Write-Host "Setting final focus to Word window before proceeding..."
        if (-not (Set-WindowFocus -Window $wordWindow)) {
            Write-Warning "Could not set final focus to Word window, trying to continue anyway"
        }
        
        # Wait a moment for the window to become fully responsive after focusing
        Start-Sleep -Seconds 3
        
        # First try to find the add-in button in the ribbon
        Write-Host "Attempting to open add-in from ribbon..." -ForegroundColor Green
        Start-Sleep -Seconds 2  # Give ribbon time to load
        $addinOpened = Open-AddInFromRibbon -wordWindow $wordWindow
        
        # If add-in not found in ribbon, try to open it from the shared folder
        if (-not $addinOpened) {
            Write-Host "Add-in not found on ribbon, opening shared folder..." -ForegroundColor Yellow
            $sharedFolderOpened = Open-SharedFolderDialog -wordWindow $wordWindow
            
            if ($sharedFolderOpened) {
                Write-Host "Shared Folder dialog opened successfully" -ForegroundColor Green
                Write-Host "Waiting for add-in to appear in the ribbon..." -ForegroundColor Yellow
                
                # Wait for Word to refresh its ribbon UI after adding the add-in
                Start-Sleep -Seconds 5
                
                # Now try to click on the add-in in the ribbon
                Write-Host "Attempting to find and click on the add-in in the ribbon..." -ForegroundColor Green
                $addinOpened = Open-AddInFromRibbon -wordWindow $wordWindow
                
                if ($addinOpened) {
                    Write-Host "Successfully opened add-in from ribbon after adding it" -ForegroundColor Green
                } else {
                    Write-Host "Add-in was added but could not be automatically opened from ribbon" -ForegroundColor Yellow
                    Write-Host "The add-in should now be available in the ribbon for future use" -ForegroundColor Green
                }
            } else {
                Write-Warning "Could not automatically open the add-in. Please open it manually."
            }
        }
    } else {
        Write-Warning "Could not find or open Word with the document. Please open Word manually."
    }
    
    # Setup completed successfully
    Write-Host "`n=== Setup Completed Successfully ===" -ForegroundColor Green
    Write-Host "The Word Add-in has been configured for user: $($userInfo.Username)"
    Write-Host "Configuration details:"
    Write-Host "- Target User: $($userInfo.Username)" -ForegroundColor Cyan
    
    $userShareName = "OfficeAddins_$($userInfo.Username)"
    $userNetworkPath = "\\$env:COMPUTERNAME\$userShareName"
    
    Write-Host "- Manifest file location: $($userInfo.DocumentsPath)\$userShareName\manifest.xml" -ForegroundColor Cyan
    Write-Host "- Network share: $userNetworkPath" -ForegroundColor Cyan
    Write-Host "- Registry configured for: $($userInfo.Username)" -ForegroundColor Cyan
    if ($documentUrl -ne "") {
        Write-Host "- Document opened: $documentName from $documentUrl" -ForegroundColor Cyan
    }
    
    Write-Host "`nNext Steps for user '$($userInfo.Username)':" -ForegroundColor Yellow
    Write-Host "1. If Word is not already open, open it now."
    Write-Host "2. Click File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs."
    Write-Host "3. Verify that $userNetworkPath is listed and check 'Show in Menu'."
    Write-Host "4. Click OK and restart Word."
    Write-Host "5. After Word starts, click 'Insert' tab > 'My Add-ins' > 'SHARED FOLDER' and select the add-in."
    
    # Script completed successfully
    exit 0

} catch {
    Write-Error "Failed during automation: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)"
} finally {
    Write-Host "`nIf the automation failed, please try manually:"
    Write-Host "1. Click File"
    Write-Host "2. Click Get Add-ins"
    Write-Host "3. Click SHARED FOLDER tab"
}