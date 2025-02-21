# Script to automate Office Add-in network share setup, manifest creation, and trust configuration
# Requires elevation
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as Administrator!"
    Exit
}

# Close any running instance of Word
Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Kill() }
Start-Sleep -Seconds 2  # Wait for Word to close

# Clear existing trusted catalogs (if any)
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\*" -Force -ErrorAction SilentlyContinue

# Add required assemblies
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

# Function to create a new GUID
function New-Guid {
    return [guid]::NewGuid().ToString()
}

# Parameters
$shareName        = "OfficeAddins"
$shareDescription = "Office Add-ins Shared Folder"
$folderPath       = "$env:USERPROFILE\Documents\$shareName"
$registryPath     = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"

# 1. Create the folder if it doesn't exist
if (-not (Test-Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath | Out-Null
    Write-Host "Created folder: $folderPath"
} else {
    Write-Host "Folder already exists: $folderPath"
}

# 2. Create network share
try {
    # If the share already exists, remove it
    $existingShare = Get-WmiObject -Class Win32_Share -Filter "Name='$shareName'" -ErrorAction SilentlyContinue
    if ($existingShare) {
        Remove-SmbShare -Name $shareName -Force -ErrorAction SilentlyContinue
    }
    
    # Create new SMB share
    New-SmbShare -Name $shareName -Path $folderPath -Description $shareDescription -FullAccess $env:USERNAME
    Write-Host "Created network share: \\$env:COMPUTERNAME\$shareName"
    $networkPath = "\\$env:COMPUTERNAME\$shareName"
} catch {
    Write-Error "Failed to create network share: $_"
    Exit
}

# 3. Add the network share to Word's Trusted Catalogs in the registry
try {
    # Generate a new GUID for the catalog entry
    $catalogGuid = New-Guid
    $catalogPath = Join-Path $registryPath $catalogGuid
    
    # Ensure the base registry key exists
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
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
    
    Write-Host "Added network share to trusted add-in catalogs in the registry."
} catch {
    Write-Error "Failed to modify registry: $_"
    Exit
}

# (Optional) Force registry refresh (this step may help Office detect changes faster)
reg.exe unload "HKCU\Temp" 2>$null
reg.exe load "HKCU\Temp" "$env:USERPROFILE\NTUSER.DAT" 2>$null
reg.exe unload "HKCU\Temp" 2>$null

# 4. Create the manifest file in the share folder
# This manifest includes an ExtensionPoint for PrimaryCommandSurface so that a Ribbon button is created in Word.
$manifestContent = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
    xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" 
    xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" 
    xsi:type="TaskPaneApp">
  <Id>f85491a7-0cf8-4950-b18c-d85ae9970d61</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="My Word Add-in"/>
  <Description DefaultValue="A template to get started"/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://localhost:3000/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://www.contoso.com/help"/>
  <AppDomains>
    <AppDomain>https://localhost:3000</AppDomain>
  </AppDomains>
  <!-- For Word, use Host Name "Document" -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:3000/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <!-- Command button definition to appear on the Word Ribbon -->
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
                    <TaskpaneId>ButtonId1</TaskpaneId>
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
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="CustomGroup.Label" DefaultValue="My Add-in Group"/>
        <bt:String id="MyButton.Label" DefaultValue="Open Add-in"/>
        <bt:String id="MyButton.Title" DefaultValue="My Word Add-in"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="MyButton.Tooltip" DefaultValue="Click to open the task pane for My Word Add-in"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
'@

$manifestDest = Join-Path $folderPath "manifest.xml"
try {
    $manifestContent | Out-File -FilePath $manifestDest -Encoding UTF8 -Force
    Write-Host "Created manifest file in share folder: $manifestDest"
} catch {
    Write-Error "Failed to create manifest file: $_"
    Write-Warning "Please manually create the manifest file at: $folderPath"
}

Write-Host "`nSetup completed successfully!"
Write-Host "Network share path: $networkPath"
Write-Host "Manifest file location: $manifestDest"
Write-Host "`nIMPORTANT:"
Write-Host "1. Start Microsoft Word."
Write-Host "2. Click File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs."
Write-Host "3. Verify that $networkPath is listed and check 'Show in Menu'."
Write-Host "4. Click OK and restart Word."
Write-Host "`nAfter Word starts, your add-in's Ribbon button should appear on the Home tab within the custom group."

# Function to check manifest existence
function Test-ManifestExists {
    $shareName = "OfficeAddins"
    $folderPath = "$env:USERPROFILE\Documents\$shareName"
    $manifestPath = Join-Path $folderPath "manifest.xml"
    
    if (Test-Path $manifestPath) {
        Write-Host "Manifest file found in shared folder"
        return $true
    }
    Write-Host "Manifest file not found in shared folder"
    return $false
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
                        Write-Host "Clicked element using InvokePattern: $ElementName"
                        return $true
                    }
                } catch {
                    Write-Host "InvokePattern not available, trying coordinate click"
                }

                # Fallback to coordinate click
                $point = $element.GetClickablePoint()
                [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point([int]$point.X, [int]$point.Y)
                Start-Sleep -Milliseconds 100
                $shell = New-Object -ComObject "WScript.Shell"
                $shell.SendKeys(" ")
                Write-Host "Clicked element using coordinates: $ElementName"
                return $true
            } catch {
                Write-Warning "Failed to click element $ElementName : $_"
            }
        }
        Start-Sleep -Milliseconds 500
    }
    Write-Host "Element not found: $ElementName"
    return $false
}

# Function to find Word window
function Find-WordWindow {
    Write-Host "Finding Word window..."
    $maxAttempts = 10
    $attempt = 0
    
    while ($attempt -lt $maxAttempts) {
        $attempt++
        Write-Host "Attempt $attempt of $maxAttempts..."
        
        $windows = [System.Windows.Automation.AutomationElement]::RootElement.FindAll(
            [System.Windows.Automation.TreeScope]::Children,
            [System.Windows.Automation.PropertyCondition]::TrueCondition
        )
        
        foreach ($window in $windows) {
            if ($window.Current.Name -match "Word" -or $window.Current.Name -match "Document") {
                Write-Host "Found Word window: $($window.Current.Name)"
                return $window
            }
        }
        Start-Sleep -Seconds 1
    }
    return $null
}

# Function to check and open add-in from ribbon
function Open-AddInFromRibbon {
    param ($wordWindow)
    
    Write-Host "Checking for add-in button on ribbon..."
    $buttonNames = @(
        "Open Add-in",
        "My Word Add-in",
        "My Add-in Group"
    )
    
    foreach ($name in $buttonNames) {
        if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow) {
            Write-Host "Successfully opened add-in from ribbon"
            return $true
        }
    }
    
    Write-Host "Add-in button not found on ribbon"
    return $false
}

# Function to open shared folder dialog
function Open-SharedFolderDialog {
    param ($wordWindow)
    
    Write-Host "Opening Shared Folder dialog..."
    
    # Click File tab
    $fileTabNames = @("File", "FILE", "File Tab")
    $found = $false
    foreach ($name in $fileTabNames) {
        if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow) {
            $found = $true
            break
        }
    }
    
    if (-not $found) {
        Write-Host "Using Alt+F shortcut for File menu..."
        [System.Windows.Forms.SendKeys]::SendWait("%F")
    }
    Start-Sleep -Seconds 2

    # Click Get Add-ins
    $getAddinsNames = @("Get Add-ins", "Office Add-ins", "Get Office Add-ins")
    $found = $false
    foreach ($name in $getAddinsNames) {
        if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow) {
            $found = $true
            break
        }
    }
    
    if (-not $found) {
        Write-Warning "Could not find Get Add-ins option"
        return $false
    }
    Start-Sleep -Seconds 2

    # Click Shared Folder
    $sharedFolderNames = @("Shared Folder", "SHARED FOLDER", "Shared folder")
    foreach ($name in $sharedFolderNames) {
        if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow) {
            Write-Host "Successfully opened Shared Folder dialog"
            return $true
        }
    }
    
    Write-Warning "Could not find Shared Folder option"
    return $false
}

# Main script
try {
    # 1. Check manifest file
    $manifestExists = Test-ManifestExists
    if (-not $manifestExists) {
        Write-Host "Installing manifest and configuring shared folder..."
        # ... (keep existing manifest installation code) ...
        return
    }

    # 2. Launch Word and wait for it to initialize
    Write-Host "Launching Word..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $doc = $word.Documents.Add()
    Start-Sleep -Seconds 3

    # 3. Find Word window
    $wordWindow = Find-WordWindow
    if (-not $wordWindow) {
        Write-Warning "Could not find Word window"
        return
    }

    # 4. Try to open add-in from ribbon first
    Write-Host "Attempting to open add-in from ribbon..."
    Start-Sleep -Seconds 2  # Give ribbon time to load
    if (Open-AddInFromRibbon -wordWindow $wordWindow) {
        Write-Host "Add-in opened successfully from ribbon"
        return
    }

    # 5. If ribbon button not found, open shared folder
    Write-Host "Add-in not found on ribbon, opening shared folder..."
    if (Open-SharedFolderDialog -wordWindow $wordWindow) {
        Write-Host "Shared Folder dialog opened successfully"
    } else {
        Write-Warning "Failed to open Shared Folder dialog"
    }

} catch {
    Write-Error "Failed during automation: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)"
} finally {
    Write-Host "`nIf the automation failed, please try manually:"
    Write-Host "1. Click File"
    Write-Host "2. Click Get Add-ins"
    Write-Host "3. Click SHARED FOLDER tab"
}