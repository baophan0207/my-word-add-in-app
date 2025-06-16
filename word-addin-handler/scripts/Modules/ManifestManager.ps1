# Office Add-in Setup - Manifest Management Module
# This module handles creation and management of manifest files

# Import required modules
. "$PSScriptRoot\..\Core\Constants.ps1"
. "$PSScriptRoot\..\Core\Logging.ps1"
. "$PSScriptRoot\..\Core\UserUtils.ps1"

# Function to generate manifest content
function New-ManifestContent {
    param(
        [hashtable]$AddinConfig = $null
    )
    
    if (-not $AddinConfig) {
        $AddinConfig = $Global:ADDIN_CONFIG
    }
    
    $manifestContent = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>$($AddinConfig.ID)</Id>
  <Version>$($AddinConfig.VERSION)</Version>
  <ProviderName>$($AddinConfig.PROVIDER)</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="$($AddinConfig.NAME)"/>
  <Description DefaultValue="A template to get started"/>
  <IconUrl DefaultValue="$($AddinConfig.BASE_URL)/assets/logo-32.png"/>
  <HighResolutionIconUrl DefaultValue="$($AddinConfig.BASE_URL)/assets/logo-64.png"/>
  <SupportUrl DefaultValue="http://www.anygenai.com/help"/>
  <AppDomains>
    <AppDomain>$($AddinConfig.BASE_URL)</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="$($AddinConfig.BASE_URL)/taskpane.html"/>
  </DefaultSettings>
  <Permissions>$($AddinConfig.PERMISSIONS)</Permissions>
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
        <bt:Image id="Icon.16x16" DefaultValue="$($AddinConfig.BASE_URL)/assets/logo-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="$($AddinConfig.BASE_URL)/assets/logo-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="$($AddinConfig.BASE_URL)/assets/logo-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="http://go.microsoft.com/fwlink/?LinkId=276812"/>
        <bt:Url id="Commands.Url" DefaultValue="$($AddinConfig.BASE_URL)/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="$($AddinConfig.BASE_URL)/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="Get started with your sample add-in!"/>
        <bt:String id="CustomGroup.Label" DefaultValue="$($AddinConfig.NAME) Group"/>
        <bt:String id="MyButton.Label" DefaultValue="$($AddinConfig.NAME)"/>
        <bt:String id="MyButton.Title" DefaultValue="$($AddinConfig.NAME)"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Your sample add-in loaded successfully. Go to the HOME tab and click the 'Show Task Pane' button to get started."/>
        <bt:String id="MyButton.Tooltip" DefaultValue="Click to open the $($AddinConfig.NAME) taskpane"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"@

    return $manifestContent
}

# Function to create manifest file for a specific user
function New-ManifestFile {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$UserInfo,
        
        [string]$ShareName = "",
        
        [string]$FolderPath = "",
        
        [hashtable]$AddinConfig = $null
    )
    
    try {
        Write-LogMessage "INFO" "Creating manifest file for user: $($UserInfo.Username)" "MANIFEST"
        Report-Progress -Status (Get-StatusCode "CREATING_MANIFEST") -Message "Creating manifest file"
        
        # Set user-specific share name if not provided
        if ($ShareName -eq "") {
            $ShareName = "OfficeAddins_$($UserInfo.Username)"
            Write-LogMessage "INFO" "Using user-specific share name: $ShareName" "MANIFEST"
        }
        
        # Set folder path if not provided
        if ($FolderPath -eq "") {
            $FolderPath = Join-Path $UserInfo.DocumentsPath $ShareName
            Write-LogMessage "INFO" "Using user's Documents folder: $FolderPath" "MANIFEST"
        }
        
        # Create the folder if it doesn't exist
        if (-not (Test-Path $FolderPath)) {
            try {
                New-Item -ItemType Directory -Path $FolderPath -Force | Out-Null
                Write-LogMessage "SUCCESS" "Created folder: $FolderPath" "MANIFEST"
            } catch {
                $errorMsg = "Failed to create folder: $FolderPath - $_"
                Write-LogMessage "ERROR" $errorMsg "MANIFEST"
                Report-Error -ErrorCode (Get-ErrorCode "FILE_SYSTEM_ACCESS_FAILED") -ErrorMessage $errorMsg -Component "MANIFEST"
                return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "FILE_SYSTEM_ACCESS_FAILED")
            }
        }
        
        # Set folder permissions for current user
        try {
            $acl = Get-Acl $FolderPath
            $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                $UserInfo.Username, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"
            )
            $acl.SetAccessRule($accessRule)
            Set-Acl -Path $FolderPath -AclObject $acl
            Write-LogMessage "SUCCESS" "Set folder permissions for user: $($UserInfo.Username)" "MANIFEST"
        } catch {
            Write-LogMessage "WARNING" "Could not set folder permissions: $_" "MANIFEST"
            # Continue anyway, might still work
        }
        
        # Generate manifest content
        $manifestContent = New-ManifestContent -AddinConfig $AddinConfig
        
        # Write manifest file
        $manifestPath = Join-Path $FolderPath (Get-Config "MANIFEST_FILENAME")
        try {
            $manifestContent | Out-File -FilePath $manifestPath -Encoding UTF8 -Force
            Write-LogMessage "SUCCESS" "Created manifest file: $manifestPath" "MANIFEST"
            
            # Verify the file was created and has content
            if ((Test-Path $manifestPath) -and ((Get-Item $manifestPath).Length -gt 0)) {
                $result = @{
                    ManifestPath = $manifestPath
                    FolderPath = $FolderPath
                    ShareName = $ShareName
                    UserInfo = $UserInfo
                }
                
                Report-Progress -Status (Get-StatusCode "CREATING_MANIFEST") -Message "Manifest file created successfully" -PercentComplete 40 -AdditionalData $result
                return New-Result -Success $true -Message "Manifest file created successfully" -Data $result
            } else {
                $errorMsg = "Manifest file was created but appears to be empty or corrupted"
                Write-LogMessage "ERROR" $errorMsg "MANIFEST"
                Report-Error -ErrorCode (Get-ErrorCode "MANIFEST_CREATION_FAILED") -ErrorMessage $errorMsg -Component "MANIFEST"
                return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "MANIFEST_CREATION_FAILED")
            }
        } catch {
            $errorMsg = "Failed to write manifest file: $_"
            Write-LogMessage "ERROR" $errorMsg "MANIFEST"
            Report-Error -ErrorCode (Get-ErrorCode "MANIFEST_CREATION_FAILED") -ErrorMessage $errorMsg -Component "MANIFEST" -Exception $_
            return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "MANIFEST_CREATION_FAILED")
        }
    }
    catch {
        $errorMsg = "Failed to create manifest file: $_"
        Write-LogMessage "ERROR" $errorMsg "MANIFEST"
        Report-Error -ErrorCode (Get-ErrorCode "MANIFEST_CREATION_FAILED") -ErrorMessage $errorMsg -Component "MANIFEST" -Exception $_
        return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "MANIFEST_CREATION_FAILED")
    }
}

# Function to check if manifest exists for a user
function Test-ManifestExists {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$UserInfo,
        
        [string]$ShareName = "",
        
        [string]$FolderPath = ""
    )
    
    try {
        Write-LogMessage "INFO" "Checking manifest file for user: $($UserInfo.Username)" "MANIFEST"
        
        # Set user-specific share name if not provided
        if ($ShareName -eq "") {
            $ShareName = "OfficeAddins_$($UserInfo.Username)"
        }
        
        # Set folder path if not provided
        if ($FolderPath -eq "") {
            $FolderPath = Join-Path $UserInfo.DocumentsPath $ShareName
        }
        
        # Check if manifest exists
        $manifestPath = Join-Path $FolderPath (Get-Config "MANIFEST_FILENAME")
        
        if (Test-Path $manifestPath) {
            # Verify the file has content
            $fileInfo = Get-Item $manifestPath
            if ($fileInfo.Length -gt 0) {
                Write-LogMessage "SUCCESS" "Manifest file found: $manifestPath" "MANIFEST"
                return @{
                    Exists = $true
                    Path = $manifestPath
                    Size = $fileInfo.Length
                    LastModified = $fileInfo.LastWriteTime
                }
            } else {
                Write-LogMessage "WARNING" "Manifest file exists but is empty: $manifestPath" "MANIFEST"
                return @{
                    Exists = $false
                    Path = $manifestPath
                    Size = 0
                    Issue = "File is empty"
                }
            }
        } else {
            Write-LogMessage "INFO" "Manifest file not found: $manifestPath" "MANIFEST"
            return @{
                Exists = $false
                Path = $manifestPath
                Issue = "File not found"
            }
        }
    }
    catch {
        Write-LogMessage "ERROR" "Error checking manifest existence: $_" "MANIFEST"
        return @{
            Exists = $false
            Path = $null
            Issue = "Error checking file: $_"
        }
    }
}

# Function to create manifests for all users on the system
function New-ManifestForAllUsers {
    param(
        [hashtable]$AddinConfig = $null,
        [switch]$SkipExisting = $true
    )
    
    try {
        Write-LogMessage "INFO" "Creating manifest files for all users" "MANIFEST"
        
        $users = Get-SystemUsers -ExcludeSystemAccounts
        $results = @()
        
        foreach ($user in $users) {
            Write-LogMessage "INFO" "Processing user: $($user.Username)" "MANIFEST"
            
            # Check if manifest already exists
            if ($SkipExisting) {
                $existsCheck = Test-ManifestExists -UserInfo $user
                if ($existsCheck.Exists) {
                    Write-LogMessage "INFO" "Manifest already exists for user $($user.Username), skipping" "MANIFEST"
                    $results += @{
                        Username = $user.Username
                        Success = $true
                        Message = "Manifest already exists"
                        Skipped = $true
                    }
                    continue
                }
            }
            
            # Create manifest for this user
            $result = New-ManifestFile -UserInfo $user -AddinConfig $AddinConfig
            $results += @{
                Username = $user.Username
                Success = $result.Success
                Message = $result.Message
                ErrorCode = $result.ErrorCode
                Data = $result.Data
                Skipped = $false
            }
            
            if ($result.Success) {
                Write-LogMessage "SUCCESS" "Created manifest for user: $($user.Username)" "MANIFEST"
            } else {
                Write-LogMessage "ERROR" "Failed to create manifest for user: $($user.Username) - $($result.Message)" "MANIFEST"
            }
        }
        
        $successCount = ($results | Where-Object { $_.Success }).Count
        $totalCount = $results.Count
        
        Write-LogMessage "SUCCESS" "Manifest creation completed: $successCount/$totalCount users" "MANIFEST"
        
        return New-Result -Success ($successCount -gt 0) -Message "Created manifests for $successCount/$totalCount users" -Data @{
            Results = $results
            SuccessCount = $successCount
            TotalCount = $totalCount
        }
    }
    catch {
        $errorMsg = "Failed to create manifests for all users: $_"
        Write-LogMessage "ERROR" $errorMsg "MANIFEST"
        return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "MANIFEST_CREATION_FAILED")
    }
}

# Note: Export-ModuleMember is not needed for .ps1 files that are dot-sourced 