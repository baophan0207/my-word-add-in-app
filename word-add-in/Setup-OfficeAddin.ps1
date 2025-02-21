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

# Function to find and click UI elements
function Find-AndClickElement {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ElementName,
        [Parameter(Mandatory=$true)]
        [System.Windows.Automation.AutomationElement]$ParentElement,
        [int]$TimeoutSeconds = 10
    )
    
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
                # Try different patterns
                # 1. Try Invoke Pattern
                try {
                    $invokePattern = $element.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                    if ($invokePattern) {
                        $invokePattern.Invoke()
                        return $true
                    }
                } catch {}

                # 2. Try Toggle Pattern
                try {
                    $togglePattern = $element.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
                    if ($togglePattern) {
                        $togglePattern.Toggle()
                        return $true
                    }
                } catch {}

                # 3. Try Expand/Collapse Pattern
                try {
                    $expandPattern = $element.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern)
                    if ($expandPattern) {
                        $expandPattern.Expand()
                        return $true
                    }
                } catch {}

                # 4. Try using SelectionItem Pattern
                try {
                    $selectionPattern = $element.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
                    if ($selectionPattern) {
                        $selectionPattern.Select()
                        return $true
                    }
                } catch {}

                # If no pattern worked, try to simulate a click
                try {
                    $clickablePoint = $element.GetClickablePoint()
                    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point([int]$clickablePoint.X, [int]$clickablePoint.Y)
                    Start-Sleep -Milliseconds 100
                    [System.Windows.Forms.SendKeys]::SendWait(" ")
                    return $true
                } catch {
                    Write-Warning "Could not click element using any method: $_"
                }
            } catch {
                Write-Warning "Error interacting with element: $_"
            }
        }
        Start-Sleep -Milliseconds 500
    }
    return $false
}

# Function to find Word window
function Find-WordWindow {
    param (
        [int]$MaxAttempts = 10,
        [int]$DelaySeconds = 1
    )
    
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        Write-Host "Attempting to find Word window (Attempt $i of $MaxAttempts)..."
        
        $processes = Get-Process | Where-Object { $_.ProcessName -eq "WINWORD" -and $_.MainWindowTitle -ne "" }
        foreach ($process in $processes) {
            $condition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ProcessIdProperty, 
                $process.Id
            )
            
            $window = [System.Windows.Automation.AutomationElement]::RootElement.FindFirst(
                [System.Windows.Automation.TreeScope]::Children,
                $condition
            )
            
            if ($window) {
                Write-Host "Found Word window: $($process.MainWindowTitle)"
                return $window
            }
        }
        
        Start-Sleep -Seconds $DelaySeconds
    }
    
    return $null
}

# Launch Word and automate UI
try {
    Write-Host "Launching Word and automating UI..."
    
    # Launch Word and create new document
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $doc = $word.Documents.Add()
    Start-Sleep -Seconds 3  # Give Word time to initialize
    
    # Find Word window
    $wordWindow = Find-WordWindow
    if (-not $wordWindow) {
        Write-Warning "Could not find Word window"
        return
    }

    # Step 1: Click File tab
    Write-Host "Clicking File tab..."
    $fileTabNames = @(
        "File",
        "FILE",
        "File Tab"
    )
    
    $found = $false
    foreach ($name in $fileTabNames) {
        if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow) {
            $found = $true
            Write-Host "Found and clicked File tab"
            break
        }
    }
    
    if (-not $found) {
        # Try using Alt+F as fallback
        Write-Host "Trying Alt+F for File menu..."
        [System.Windows.Forms.SendKeys]::SendWait("%F")
        Start-Sleep -Seconds 1
    }
    Start-Sleep -Seconds 2  # Wait for File menu to open

    # Step 2: Click Get Add-ins
    Write-Host "Clicking Get Add-ins..."
    $getAddinsNames = @(
        "Get Add-ins",
        "Office Add-ins",
        "Get Office Add-ins"
    )
    
    $found = $false
    foreach ($name in $getAddinsNames) {
        if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow) {
            $found = $true
            Write-Host "Found and clicked Get Add-ins"
            break
        }
    }
    
    if (-not $found) {
        Write-Warning "Could not find Get Add-ins option"
        return
    }
    Start-Sleep -Seconds 2  # Wait for dialog to open

    # Step 3: Click Shared Folder
    Write-Host "Clicking Shared Folder..."
    $sharedFolderNames = @(
        "Shared Folder",
        "SHARED FOLDER",
        "Shared folder"
    )
    
    $found = $false
    foreach ($name in $sharedFolderNames) {
        if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow) {
            $found = $true
            Write-Host "Found and clicked Shared Folder"
            break
        }
    }
    
    if (-not $found) {
        Write-Warning "Could not find Shared Folder option"
        return
    }
    Start-Sleep -Seconds 2

    # Step 4: Click the add-in
    Write-Host "Clicking add-in..."
    $addinNames = @(
        "My Word Add-in",
        "My Word Add",
        "My Word Add..."  # For truncated names
    )
    
    $found = $false
    foreach ($name in $addinNames) {
        if (Find-AndClickElement -ElementName $name -ParentElement $wordWindow) {
            $found = $true
            Write-Host "Found and clicked add-in"
            break
        }
    }
    
    if (-not $found) {
        Write-Warning "Could not find add-in in the list"
        return
    }

    Write-Host "UI automation completed successfully"
    
} catch {
    Write-Error "Failed during UI automation: $_"
    Write-Host "Stack Trace: $($_.ScriptStackTrace)"
} finally {
    Write-Host "`nIf the automation failed, please try manually:"
    Write-Host "1. Click File"
    Write-Host "2. Click Get Add-ins"
    Write-Host "3. Click Shared Folder"
    Write-Host "4. Select 'My Word Add-in'"
}

# Alternative approach using SendKeys as fallback
if (-not $found) {
    Write-Host "Trying alternative approach with keyboard shortcuts..."
    try {
        # Alt+F for File menu
        [System.Windows.Forms.SendKeys]::SendWait("%F")
        Start-Sleep -Seconds 1
        
        # Navigate to Get Add-ins (might need adjusting based on your Word version)
        [System.Windows.Forms.SendKeys]::SendWait("A")
        Start-Sleep -Seconds 2
        
        # Navigate to Shared Folder
        [System.Windows.Forms.SendKeys]::SendWait("S")
        Start-Sleep -Seconds 1
        
        Write-Host "Keyboard navigation completed"
    } catch {
        Write-Error "Failed to send keystrokes: $_"
    }
}
