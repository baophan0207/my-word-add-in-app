# Script to automate Office Add-in network share setup, manifest creation, and trust configuration
# Requires elevation
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as Administrator!"
    Exit
}

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

function Install-Addin {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]$ShareName = "OfficeAddins",
        
        [Parameter()]
        [string]$ShareDescription = "Office Add-ins Shared Folder",
        
        [Parameter()]
        [string]$FolderPath = "$env:USERPROFILE\Documents\$ShareName",
        
        [Parameter()]
        [string]$RegistryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
    )

    # 1. Create the folder if it doesn't exist
    if (-not (Test-Path $FolderPath)) {
        New-Item -ItemType Directory -Path $FolderPath | Out-Null
        Write-Host "Created folder: $FolderPath"
    } else {
        Write-Host "Folder already exists: $FolderPath"
    }

    # 2. Create network share
    try {
        # If the share already exists, remove it
        $existingShare = Get-WmiObject -Class Win32_Share -Filter "Name='$ShareName'" -ErrorAction SilentlyContinue
        if ($existingShare) {
            Remove-SmbShare -Name $ShareName -Force -ErrorAction SilentlyContinue
        }
        
        # Create new SMB share
        New-SmbShare -Name $ShareName -Path $FolderPath -Description $ShareDescription -FullAccess $env:USERNAME
        Write-Host "Created network share: \\$env:COMPUTERNAME\$ShareName"
        $networkPath = "\\$env:COMPUTERNAME\$ShareName"
    } catch {
        Write-Error "Failed to create network share: $_"
        return $false
    }

    # 3. Add the network share to Word's Trusted Catalogs in the registry
    try {
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
        
        Write-Host "Added network share to trusted add-in catalogs in the registry."
    } catch {
        Write-Error "Failed to modify registry: $_"
        return $false
    }

    # (Optional) Force registry refresh (this step may help Office detect changes faster)
    reg.exe unload "HKCU\Temp" 2>$null
    reg.exe load "HKCU\Temp" "$env:USERPROFILE\NTUSER.DAT" 2>$null
    reg.exe unload "HKCU\Temp" 2>$null

    # 4. Create the manifest file in the share folder
    # This manifest includes an ExtensionPoint for PrimaryCommandSurface so that a Ribbon button is created in Word.
    $manifestContent = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>f85491a7-0cf8-4950-b18c-d85ae9970d61</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="My Word Add-in"/>
  <Description DefaultValue="A template to get started"/>
  <IconUrl DefaultValue="https://10.100.100.71:3002/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="https://10.100.100.71:3002/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://www.contoso.com/help"/>
  <AppDomains>
    <AppDomain>https://10.100.100.71:3002</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://10.100.100.71:3002/taskpane.html"/>
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
        <bt:Image id="Icon.16x16" DefaultValue="https://10.100.100.71:3002/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://10.100.100.71:3002/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://10.100.100.71:3002/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="https://10.100.100.71:3002/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://10.100.100.71:3002/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Get started with your sample add-in!"/>
        <bt:String id="CustomGroup.Label" DefaultValue="My Add-in Group"/>
        <bt:String id="MyButton.Label" DefaultValue="Open Add-in"/>
        <bt:String id="MyButton.Title" DefaultValue="My Word Add-in"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Your sample add-in loaded successfully. Go to the HOME tab and click the 'Show Task Pane' button to get started."/>
        <bt:String id="MyButton.Tooltip" DefaultValue="Click to open the My Word Add-in taskpane"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
'@

    $manifestDest = Join-Path $FolderPath "manifest.xml"
    try {
        $manifestContent | Out-File -FilePath $manifestDest -Encoding UTF8 -Force
        Write-Host "Created manifest file in share folder: $manifestDest"
    } catch {
        Write-Error "Failed to create manifest file: $_"
        Write-Warning "Please manually create the manifest file at: $FolderPath"
        return $false
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

    # Return success and the network path for potential further use
    return @{
        Success = $true
        NetworkPath = $networkPath
        ManifestPath = $manifestDest
    }
}

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

# Function to find Word window with a specific document name
function Find-WordWindowWithDocument {
    Write-Host "Finding already open Word document window..."
    $maxAttempts = 15
    $attempt = 0
    
    while ($attempt -lt $maxAttempts) {
        $attempt++
        Write-Host "Attempt $attempt of $maxAttempts..."
        
        # Get Word processes to verify we're working with actual Word windows
        $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
        
        if ($wordProcesses.Count -eq 0) {
            Write-Host "No Word processes found, waiting..."
            Start-Sleep -Seconds 1
            continue
        }

        Write-Host "Found $($wordProcesses.Count) Word process(es)"
        
        # Get all top-level windows
        $windows = [System.Windows.Automation.AutomationElement]::RootElement.FindAll(
            [System.Windows.Automation.TreeScope]::Children,
            [System.Windows.Automation.PropertyCondition]::TrueCondition
        )
        
        foreach ($window in $windows) {
            $windowName = $window.Current.Name
            $className = $window.Current.ClassName
            
            # Skip windows that are clearly not Word (like PowerShell or VSCode windows)
            if ($windowName -match "\.ps1" -or $windowName -match "PowerShell" -or 
                $windowName -match "Visual Studio Code") {
                continue
            }
            
            # Check for proper Word window class and process ID
            $isWordWindow = $false
            
            # Look for Word-specific class names (OpusApp is Word's main window class)
            if ($className -eq "OpusApp") {
                $isWordWindow = $true
            }
            else {
                # Try to verify by checking if the window's process is Word
                try {
                    # Get process ID from window handle
                    $hwnd = $window.Current.NativeWindowHandle
                    if ($hwnd -ne 0) {
                        $processId = 0
                        $null = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
                        $threadId = [Win32Functions.Win32Api]::GetWindowThreadProcessId($hwnd, [ref]$processId)
                        
                        # Check if process ID matches any Word process
                        if ($wordProcesses.Id -contains $processId) {
                            $isWordWindow = $true
                        }
                    }
                }
                catch {
                    Write-Host "Error getting process ID: $_"
                }
            }
            
            if ($isWordWindow) {
                Write-Host "Found Word window: $windowName (Class: $className)"
                
                # Extract document name (everything before " - Word" in the window title)
                $documentName = $null
                if ($windowName -match "(.+) - Word") {
                    $documentName = $matches[1]
                    Write-Host "Document name: $documentName"
                }
                
                # Check if it's in Protected View by looking for the Enable Editing button
                $enableEditingCondition = New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::NameProperty, 
                    "Enable Editing"
                )
                
                $enableEditingButton = $window.FindFirst(
                    [System.Windows.Automation.TreeScope]::Descendants,
                    $enableEditingCondition
                )
                
                if ($enableEditingButton) {
                    Write-Host "Document is in Protected View, will enable editing"
                    return @{
                        Window = $window
                        EnableEditingButton = $enableEditingButton
                        DocumentName = $documentName
                    }
                } else {
                    # This could be the document already in editing mode
                    Write-Host "Found Word window in editing mode"
                    return @{
                        Window = $window
                        EnableEditingButton = $null
                        DocumentName = $documentName
                    }
                }
            }
        }
        Start-Sleep -Seconds 1
    }
    return $null
}

# Function to find Word window with a specific document name
function Find-WordWindowWithName {
    param (
        [Parameter(Mandatory=$true)]
        [string]$DocumentName
    )
    
    Write-Host "Looking for Word window with document name: $DocumentName"
    $maxAttempts = 10
    $attempt = 0
    
    while ($attempt -lt $maxAttempts) {
        $attempt++
        
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
            
            # For edited documents, the window title might now be "DocumentName - Word"
            # or might include additional information if it was just saved
            $doesMatch = $false
            if ($windowName -eq "$DocumentName - Word" -or 
                $windowName -match "$DocumentName\.docx? - Word" -or
                $windowName -match "$DocumentName \[.*\] - Word") {
                $doesMatch = $true
            }
            
            # Check if it's a Word window by class name or process
            if ($doesMatch) {
                if ($className -eq "OpusApp" -or (IsWordProcess -window $window)) {
                    Write-Host "Found matching document window: $windowName"
                    return $window
                }
            }
        }
        
        Write-Host "Document window not found on attempt $attempt, waiting..."
        Start-Sleep -Seconds 1
    }
    
    Write-Host "Could not find document window with name: $DocumentName"
    return $null
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
            $threadId = [Win32Functions.Win32Api]::GetWindowThreadProcessId($hwnd, [ref]$processId)
            
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

# Function to enable editing on a protected document
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
                    Write-Host "Clicked Enable Editing button using InvokePattern"
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
            
            Write-Host "Clicked Enable Editing button using coordinates"
            Start-Sleep -Seconds 2  # Wait for document to exit Protected View
            return $true
        } catch {
            Write-Warning "Failed to click Enable Editing button: $_"
            return $false
        }
    } else {
        Write-Host "No Enable Editing button found, document may already be in edit mode"
        return $true
    }
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
    Start-Sleep -Seconds 3  # Wait for dialog to open

    # Find Office Add-ins dialog
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
    Write-Host "Found Office Add-ins dialog"

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
        try {
            $invokePattern = $sharedFolderTab.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
            if ($invokePattern) {
                $invokePattern.Invoke()
                Write-Host "Clicked SHARED FOLDER tab"
            }
        } catch {
            Write-Host "Using alternative method to click SHARED FOLDER tab"
            $point = $sharedFolderTab.GetClickablePoint()
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point([int]$point.X, [int]$point.Y)
            Start-Sleep -Milliseconds 100
            $shell = New-Object -ComObject "WScript.Shell"
            $shell.SendKeys(" ")
        }
    } else {
        Write-Warning "Could not find SHARED FOLDER tab"
        return $false
    }
    Start-Sleep -Seconds 2  # Wait for tab content to load

    # After clicking SHARED FOLDER tab and waiting for content to load
    Start-Sleep -Seconds 2

    Write-Host "Looking for My Word Add-in in the dialog..."
    
    # Try multiple approaches to find the add-in element
    $addInFound = $false
    
    # Approach 1: Direct name search
    $nameCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty, 
        "My Word Add-in"
    )
    
    $addInElement = $addinDialog.FindFirst(
        [System.Windows.Automation.TreeScope]::Descendants,
        $nameCondition
    )

    # Approach 2: Look for list item containing both "My Word Add-in" and "Contoso"
    if (-not $addInElement) {
        Write-Host "Trying to find add-in by list item..."
        
        # Find all list items
        $listItemCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::ListItem
        )
        
        $listItems = $addinDialog.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            $listItemCondition
        )

        $targetItem = $null
        foreach ($item in $listItems) {
            if ($item.Current.Name -match "My Word Add-in") {
                $targetItem = $item
                Write-Host "Found target list item: $($item.Current.Name)"
                break
            }
        }

        if ($targetItem) {
            Write-Host "Found target item, attempting to select..."
            try {
                # Try to find the list container first
                $listCondition = New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                    [System.Windows.Automation.ControlType]::List
                )
                
                $listContainer = $addinDialog.FindFirst(
                    [System.Windows.Automation.TreeScope]::Descendants,
                    $listCondition
                )

                if ($listContainer) {
                    Write-Host "Found list container, attempting to select item..."
                    
                    # Try ExpandCollapsePattern if available
                    try {
                        $expandPattern = $listContainer.GetCurrentPattern(
                            [System.Windows.Automation.ExpandCollapsePattern]::Pattern
                        )
                        if ($expandPattern) {
                            $expandPattern.Expand()
                            Write-Host "Expanded list container"
                            Start-Sleep -Milliseconds 500
                        }
                    } catch {
                        Write-Host "ExpandCollapsePattern not available or failed: $_"
                    }

                    # Get the bounding rectangle of the target item
                    $boundingRect = $targetItem.Current.BoundingRectangle
                    
                    # Calculate click position (center of the item)
                    $clickX = $boundingRect.X + ($boundingRect.Width / 2)
                    $clickY = $boundingRect.Y + ($boundingRect.Height / 2)
                    
                    # Move mouse and click
                    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(
                        [int]$clickX, 
                        [int]$clickY
                    )
                    Start-Sleep -Milliseconds 200
                    
                    # Check if type already exists before adding it
                    if (-not ([System.Management.Automation.PSTypeName]'Win32Functions.Win32MouseEventNew').Type) {
                        $signature = @'
                        [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
                        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
'@
                        Add-Type -MemberDefinition $signature -Name "Win32MouseEventNew" -Namespace Win32Functions
                    }
                    
                    # Get the type and create mouse click
                    $SendMouseClick = [Win32Functions.Win32MouseEventNew]
                    
                    # Mouse click down
                    $SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0)
                    Start-Sleep -Milliseconds 100
                    # Mouse click up
                    $SendMouseClick::mouse_event(0x00000004, 0, 0, 0, 0)
                    
                    Write-Host "Clicked on add-in using calculated center position"
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
                        Write-Host "Found Add button, attempting to click..."
                        try {
                            $invokePattern = $addButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                            if ($invokePattern) {
                                $invokePattern.Invoke()
                                Write-Host "Clicked Add button using InvokePattern"
                                Start-Sleep -Seconds 5
                                
                                # Wait and try to find the ribbon button
                                Write-Host "Waiting for add-in button to appear on ribbon..."
                                $maxAttempts = 20
                                $attempt = 0
                                $buttonFound = $false

                                while ($attempt -lt $maxAttempts -and -not $buttonFound) {
                                    $attempt++
                                    Write-Host "Attempt $attempt of $maxAttempts to find ribbon button..."

                                    # Try different possible button names
                                    $buttonNames = @(
                                        "Open Add-in",
                                        "My Word Add-in",
                                        "My Add-in Group"
                                    )

                                    foreach ($name in $buttonNames) {
                                        $buttonCondition = New-Object System.Windows.Automation.PropertyCondition(
                                            [System.Windows.Automation.AutomationElement]::NameProperty, 
                                            $name
                                        )
                                        
                                        $ribbonButton = $wordWindow.FindFirst(
                                            [System.Windows.Automation.TreeScope]::Descendants,
                                            $buttonCondition
                                        )

                                        if ($ribbonButton) {
                                            Write-Host "Found ribbon button: $name"
                                            Start-Sleep -Seconds 2
                                            try {
                                                $invokePattern = $ribbonButton.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                                                if ($invokePattern) {
                                                    $invokePattern.Invoke()
                                                    Write-Host "Clicked ribbon button using InvokePattern"
                                                    $buttonFound = $true
                                                    break
                                                }
                                            } catch {
                                                Write-Host "InvokePattern not available, trying coordinate click"
                                                $point = $ribbonButton.GetClickablePoint()
                                                [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(
                                                    [int]$point.X, 
                                                    [int]$point.Y
                                                )
                                                Start-Sleep -Milliseconds 500
                                                $SendMouseClick::mouse_event(0x00000002, 0, 0, 0, 0)
                                                Start-Sleep -Milliseconds 200
                                                $SendMouseClick::mouse_event(0x00000004, 0, 0, 0, 0)
                                                Write-Host "Clicked ribbon button using coordinates"
                                                $buttonFound = $true
                                                break
                                            }
                                        }
                                    }

                                    if (-not $buttonFound) {
                                        Start-Sleep -Seconds 3
                                    }
                                }

                                if ($buttonFound) {
                                    Write-Host "Successfully added and opened add-in"
                                    Start-Sleep -Seconds 2
                                    return $true
                                } else {
                                    Write-Warning "Add-in installed but couldn't find ribbon button"
                                    return $false
                                }
                            }
                        } catch {
                            Write-Warning "Failed to click Add button: $_"
                            return $false
                        }
                    } else {
                        Write-Warning "Could not find Add button"
                        return $false
                    }
                } else {
                    Write-Warning "Could not find list container"
                    return $false
                }
            } catch {
                Write-Warning "Failed to select add-in: $_"
                return $false
            }
        } else {
            Write-Warning "Could not find target list item"
            return $false
        }
    } else {
        Write-Warning "Could not find My Word Add-in element"
        return $false
    }

    return $false
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
            
            if ($result) {
                Write-Host "Successfully set focus to Word window"
            } else {
                Write-Warning "SetForegroundWindow API call failed"
            }
            
            # Additional focus attempt using UI Automation
            try {
                $window.SetFocus()
                Write-Host "Set focus using UI Automation"
            } catch {
                Write-Host "Could not set focus using UI Automation: $_"
            }
            
            # Wait for window to become responsive
            Start-Sleep -Seconds 2
            
            return $result
        } else {
            Write-Warning "Invalid window handle"
            return $false
        }
    } catch {
        Write-Warning "Failed to set window focus: $_"
        return $false
    }
}

# Main script
try {
    # 1. Check manifest file
    $manifestExists = Test-ManifestExists
    if (-not $manifestExists) {
        Write-Host "Installing manifest and configuring shared folder..."

        # Or with custom parameters
        $result = Install-Addin -ShareName "WordAddins" -ShareDescription "My Word Add-ins"

        if ($result.Success) {
            Write-Host "Add-in installation succeeded"
            # Additional actions using $result.NetworkPath or $result.ManifestPath
        } else {
            Write-Warning "Add-in installation failed"
            return
        }
    }

    # 2. Find already open Word document window
    Write-Host "Looking for the already opened Word document..."
    
    # Check we have Word processes running first
    $wordProcesses = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue
    if ($wordProcesses.Count -eq 0) {
        Write-Warning "No Word processes found. Please ensure Word is running with a document open."
        return
    }
    
    Write-Host "Found $($wordProcesses.Count) Word process(es)"
    
    # Find Word windows
    $wordInfo = Find-WordWindowWithDocument
    if (-not $wordInfo) {
        Write-Warning "Could not find any open Word documents"
        return
    }
    
    $wordWindow = $wordInfo.Window
    $documentName = $wordInfo.DocumentName
    
    Write-Host "Found document: $documentName"
    
    # 3. Set focus to Word window first (before enabling editing)
    Write-Host "Setting focus to Word window before enabling editing..."
    if (-not (Set-WindowFocus -Window $wordWindow)) {
        Write-Warning "Could not set focus to Word window before enabling editing"
        # Continue anyway as we might still be able to enable editing
    }
    
    Start-Sleep -Seconds 2  # Wait for window to be fully focused
    
    # 4. Enable editing if document is in Protected View
    if ($wordInfo.EnableEditingButton) {
        Write-Host "Document is in Protected View, enabling editing..."
        if (-not (Enable-DocumentEditing -enableEditingButton $wordInfo.EnableEditingButton)) {
            Write-Warning "Failed to enable editing on document"
            # Continue anyway as the document might still be usable
        }
        Start-Sleep -Seconds 2  # Wait for document to exit Protected View
        
        # 5. After enabling editing, find the document window again
        if ($documentName) {
            Write-Host "Re-locating document after enabling editing: $documentName"
            # Wait a bit longer to ensure Word has updated the window state
            Start-Sleep -Seconds 3
            
            $updatedWindow = Find-WordWindowWithName -DocumentName $documentName
            
            if ($updatedWindow) {
                Write-Host "Re-located document window after enabling editing"
                $wordWindow = $updatedWindow
            } else {
                Write-Warning "Could not re-locate document window by name after enabling editing"
                
                # As a fallback, try to find any Word window
                Write-Host "Trying to find any Word window as fallback..."
                $fallbackInfo = Find-WordWindowWithDocument
                if ($fallbackInfo) {
                    Write-Host "Found fallback Word window: $($fallbackInfo.DocumentName)"
                    $wordWindow = $fallbackInfo.Window
                } else {
                    Write-Warning "Could not find any Word window after enabling editing, using original window"
                }
            }
        }
    }
    
    # 6. Ensure the Word window has focus before proceeding
    Write-Host "Setting final focus to Word window before proceeding..."
    if (-not (Set-WindowFocus -Window $wordWindow)) {
        Write-Warning "Could not set final focus to Word window, trying to continue anyway"
    }
    
    # Wait a moment for the window to become fully responsive after focusing
    Start-Sleep -Seconds 3

    # 7. Try to open add-in from ribbon first
    Write-Host "Attempting to open add-in from ribbon..."
    Start-Sleep -Seconds 2  # Give ribbon time to load
    if (Open-AddInFromRibbon -wordWindow $wordWindow) {
        Write-Host "Add-in opened successfully from ribbon"
        return
    }

    # 8. If ribbon button not found, open shared folder
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