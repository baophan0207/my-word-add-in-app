# Script to automate Office Add-in network share setup and trust configuration

# Add Windows Forms assembly for SendKeys
Add-Type -AssemblyName System.Windows.Forms

# Requires elevation
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))  
{  
    Write-Warning "Please run this script as Administrator!"
    Exit
}

# Force close Word if it's running
Get-Process | Where-Object {$_.ProcessName -eq "WINWORD"} | ForEach-Object { $_.Kill() }
Start-Sleep -Seconds 2  # Wait for Word to fully close

# Clear existing trusted catalogs
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\*" -Force -ErrorAction SilentlyContinue

# Function to create a new GUID
function New-Guid {
    return [guid]::NewGuid().ToString()
}

# Parameters
$shareName = "OfficeAddins"
$shareDescription = "Office Add-ins Shared Folder"
$folderPath = "$env:USERPROFILE\Documents\$shareName"
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"

# 1. Create the folder if it doesn't exist
if (-not (Test-Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath
    Write-Host "Created folder: $folderPath"
}

# 2. Create network share
try {
    # Remove share if it already exists
    if (Get-WmiObject -Class Win32_Share -Filter "Name='$shareName'") {
        Remove-SmbShare -Name $shareName -Force
    }
    
    # Create new share
    New-SmbShare -Name $shareName -Path $folderPath -Description $shareDescription -FullAccess $env:USERNAME
    Write-Host "Created network share: \\$env:COMPUTERNAME\$shareName"
    
    # Get the full network path
    $networkPath = "\\$env:COMPUTERNAME\$shareName"
} catch {
    Write-Error "Failed to create network share: $_"
    Exit
}

# 3. Add to trusted catalogs in registry
try {
    # Create a new GUID for the catalog
    $catalogGuid = New-Guid
    $catalogPath = Join-Path $registryPath $catalogGuid
    
    # Create registry keys
    if (-not (Test-Path $registryPath)) {
        New-Item -Path $registryPath -Force | Out-Null
    }
    
    New-Item -Path $catalogPath -Force | Out-Null
    
    # Add registry values
    New-ItemProperty -Path $catalogPath -Name "Id" -Value $catalogGuid -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $catalogPath -Name "Url" -Value $networkPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $catalogPath -Name "Flags" -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $catalogPath -Name "Type" -Value 2 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $catalogPath -Name "CatalogVersion" -Value 2 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $catalogPath -Name "SkipCatalogUpdate" -Value 0 -PropertyType DWord -Force | Out-Null
    
    Write-Host "Added catalog to trusted locations in registry"
} catch {
    Write-Error "Failed to modify registry: $_"
    Exit
}

# 4. Create manifest file in the share folder
$manifestContent = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
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
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
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

$manifestDest = Join-Path $folderPath "manifest.xml"
try {
    $manifestContent | Out-File -FilePath $manifestDest -Encoding UTF8 -Force
    Write-Host "Created manifest file in share folder"
} catch {
    Write-Error "Failed to create manifest file: $_"
    Write-Warning "Please manually create manifest file at: $folderPath"
}

# 5. Force registry refresh
reg.exe unload "HKCU\Temp" 2>$null
reg.exe load "HKCU\Temp" "$env:USERPROFILE\NTUSER.DAT" 2>$null
reg.exe unload "HKCU\Temp" 2>$null

# 6. Configure Word settings directly through registry
try {
    # Enable Developer tab
    $devModeKey = "HKCU:\Software\Microsoft\Office\16.0\Word\Options"
    New-ItemProperty -Path $devModeKey -Name "DeveloperTools" -Value 1 -PropertyType DWORD -Force | Out-Null
    
    # Set trusted location
    $trustKey = "HKCU:\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations"
    New-Item -Path "$trustKey\$shareName" -Force | Out-Null
    New-ItemProperty -Path "$trustKey\$shareName" -Name "Path" -Value $networkPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path "$trustKey\$shareName" -Name "Date" -Value (Get-Date).ToFileTime() -PropertyType QWORD -Force | Out-Null
    New-ItemProperty -Path "$trustKey\$shareName" -Name "Description" -Value "Auto-configured trusted location" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path "$trustKey\$shareName" -Name "AllowSubFolders" -Value 1 -PropertyType DWORD -Force | Out-Null

    Write-Host "Word settings configured successfully"
    
    # Launch Word
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    
    # Create a new document
    $doc = $word.Documents.Add()
    
    # Configure COM Add-in settings
    $word.COMAddIns | ForEach-Object {
        if ($_.Description -match $shareName) {
            $_.Connect = $true
        }
    }
    
    Write-Host "Word launched and add-in configured"
    
    # Display instructions for manual steps if needed
    Write-Host "`nIf the add-in is not visible, please follow these steps:"
    Write-Host "1. Click 'Insert' tab"
    Write-Host "2. Click 'My Add-ins'"
    Write-Host "3. Look for 'Shared Folder' in the dropdown"
    Write-Host "4. Select your add-in from the list"
    
} catch {
    Write-Error "Failed to configure Word: $_"
    Write-Host "Please configure the add-in manually through Word's UI"
}