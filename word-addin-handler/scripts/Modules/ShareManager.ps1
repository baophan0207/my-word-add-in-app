# Office Add-in Setup - Network Share and Trust Management Module
# This module handles network share creation and trust center configuration

# Import required modules
. "$PSScriptRoot\..\Core\Constants.ps1"
. "$PSScriptRoot\..\Core\Logging.ps1"
. "$PSScriptRoot\..\Core\UserUtils.ps1"

# Function to create network share for user
function New-UserNetworkShare {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$UserInfo,
        
        [string]$ShareName = "",
        
        [string]$FolderPath = "",
        
        [string]$ShareDescription = ""
    )
    
    try {
        Write-LogMessage "INFO" "Creating network share for user: $($UserInfo.Username)" "SHARE"
        Report-Progress -Status (Get-StatusCode "CONFIGURING_SHARE") -Message "Creating network share"
        
        # Set user-specific share name if not provided
        if ($ShareName -eq "") {
            $ShareName = "OfficeAddins_$($UserInfo.Username)"
            Write-LogMessage "INFO" "Using user-specific share name: $ShareName" "SHARE"
        }
        
        # Set folder path if not provided
        if ($FolderPath -eq "") {
            $FolderPath = Join-Path $UserInfo.DocumentsPath $ShareName
            Write-LogMessage "INFO" "Using user's Documents folder: $FolderPath" "SHARE"
        }
        
        # Set description if not provided
        if ($ShareDescription -eq "") {
            $ShareDescription = (Get-Config "DEFAULT_SHARE_DESCRIPTION") + " for $($UserInfo.Username)"
        }
        
        # Check if the share already exists
        $existingShare = Get-WmiObject -Class Win32_Share -Filter "Name='$ShareName'" -ErrorAction SilentlyContinue
        
        if ($existingShare) {
            Write-LogMessage "INFO" "Network share already exists, removing it first: $ShareName" "SHARE"
            try {
                Remove-SmbShare -Name $ShareName -Force -ErrorAction SilentlyContinue
                Write-LogMessage "SUCCESS" "Removed existing network share: $ShareName" "SHARE"
            } catch {
                Write-LogMessage "WARNING" "Could not remove existing share: $_" "SHARE"
                # Continue anyway, might still work
            }
        }
        
        # Create the folder if it doesn't exist
        if (-not (Test-Path $FolderPath)) {
            try {
                New-Item -ItemType Directory -Path $FolderPath -Force | Out-Null
                Write-LogMessage "SUCCESS" "Created folder: $FolderPath" "SHARE"
            } catch {
                $errorMsg = "Failed to create folder: $FolderPath - $_"
                Write-LogMessage "ERROR" $errorMsg "SHARE"
                Report-Error -ErrorCode (Get-ErrorCode "NETWORK_SHARE_FAILED") -ErrorMessage $errorMsg -Component "SHARE"
                return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "NETWORK_SHARE_FAILED")
            }
        }
        
        # Determine users who should have access to the share
        $shareUsers = @("Administrators")
        $isSystemAccount = $UserInfo.Username -in @("systemprofile", "LocalService", "NetworkService", "SYSTEM")
        
        if (-not $isSystemAccount -and $UserInfo.Username -ne "Administrator") {
            $shareUsers += $UserInfo.Username
        }
        
        # Create new SMB share
        try {
            if ($isSystemAccount) {
                # For system accounts, create share with basic permissions and handle warnings
                Write-LogMessage "INFO" "Creating share for system account: $($UserInfo.Username)" "SHARE"
                New-SmbShare -Name $ShareName -Path $FolderPath -Description $ShareDescription -FullAccess "Administrators" -ErrorAction Continue
            } else {
                # For regular users, include them in the access list
                New-SmbShare -Name $ShareName -Path $FolderPath -Description $ShareDescription -FullAccess $shareUsers
            }
            Write-LogMessage "SUCCESS" "Created SMB share: $ShareName at $FolderPath" "SHARE"
        } catch {
            # Check if the error is just about account mapping (common with system accounts)
            if ($_.Exception.Message -like "*No mapping between account names and security IDs*") {
                Write-LogMessage "WARNING" "Account mapping warning for system account (share still created): $($UserInfo.Username)" "SHARE"
                # The share was likely still created, so continue
            } else {
                $errorMsg = "Failed to create SMB share: $_"
                Write-LogMessage "ERROR" $errorMsg "SHARE"
                Report-Error -ErrorCode (Get-ErrorCode "NETWORK_SHARE_FAILED") -ErrorMessage $errorMsg -Component "SHARE" -Exception $_
                return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "NETWORK_SHARE_FAILED")
            }
        }
        
        # Set NTFS permissions to allow Everyone read access (needed for Office)
        try {
            $acl = Get-Acl $FolderPath
            
            # Add Everyone with ReadAndExecute permissions
            $everyoneRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                "Everyone", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow"
            )
            $acl.SetAccessRule($everyoneRule)
            
            # For regular users (not system accounts), ensure the user has full control
            if (-not $isSystemAccount) {
                try {
                    $userRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                        $UserInfo.Username, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"
                    )
                    $acl.SetAccessRule($userRule)
                    Write-LogMessage "SUCCESS" "Added full control permissions for user: $($UserInfo.Username)" "SHARE"
                } catch {
                    Write-LogMessage "WARNING" "Could not set user-specific permissions for $($UserInfo.Username): $_" "SHARE"
                    # Continue with Everyone permissions
                }
            } else {
                Write-LogMessage "INFO" "Skipping user-specific permissions for system account: $($UserInfo.Username)" "SHARE"
            }
            
            Set-Acl -Path $FolderPath -AclObject $acl
            Write-LogMessage "SUCCESS" "Set NTFS permissions for share folder" "SHARE"
        } catch {
            if ($_.Exception.Message -like "*Some or all identity references could not be translated*") {
                Write-LogMessage "WARNING" "Could not translate identity references for NTFS permissions (common with system accounts): $($UserInfo.Username)" "SHARE"
            } else {
                Write-LogMessage "WARNING" "Could not set NTFS permissions: $_" "SHARE"
            }
            # Continue anyway, share might still work
        }
        
        $networkPath = "\\$env:COMPUTERNAME\$ShareName"
        $result = @{
            ShareName = $ShareName
            FolderPath = $FolderPath
            NetworkPath = $networkPath
            ShareDescription = $ShareDescription
            UserInfo = $UserInfo
        }
        
        Write-LogMessage "SUCCESS" "Network share created successfully: $networkPath" "SHARE"
        Report-Progress -Status (Get-StatusCode "CONFIGURING_SHARE") -Message "Network share created" -PercentComplete 60 -AdditionalData $result
        
        return New-Result -Success $true -Message "Network share created successfully" -Data $result
    }
    catch {
        $errorMsg = "Failed to create network share: $_"
        Write-LogMessage "ERROR" $errorMsg "SHARE"
        Report-Error -ErrorCode (Get-ErrorCode "NETWORK_SHARE_FAILED") -ErrorMessage $errorMsg -Component "SHARE" -Exception $_
        return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "NETWORK_SHARE_FAILED")
    }
}

# Function to test if network share exists
function Test-NetworkShareExists {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$UserInfo,
        
        [string]$ShareName = ""
    )
    
    try {
        Write-LogMessage "INFO" "Checking network share for user: $($UserInfo.Username)" "SHARE"
        
        if ($ShareName -eq "") {
            $ShareName = "OfficeAddins_$($UserInfo.Username)"
        }
        
        # Check if the share exists
        $existingShare = Get-WmiObject -Class Win32_Share -Filter "Name='$ShareName'" -ErrorAction SilentlyContinue
        
        if ($existingShare) {
            $networkPath = "\\$env:COMPUTERNAME\$ShareName"
            Write-LogMessage "SUCCESS" "Network share exists: $networkPath" "SHARE"
            
            return @{
                Exists = $true
                ShareName = $ShareName
                NetworkPath = $networkPath
                LocalPath = $existingShare.Path
                Description = $existingShare.Description
            }
        } else {
            Write-LogMessage "INFO" "Network share does not exist: $ShareName" "SHARE"
            return @{
                Exists = $false
                ShareName = $ShareName
                NetworkPath = "\\$env:COMPUTERNAME\$ShareName"
            }
        }
    }
    catch {
        Write-LogMessage "ERROR" "Error checking network share: $_" "SHARE"
        return @{
            Exists = $false
            ShareName = $ShareName
            Error = $_
        }
    }
}

# Helper function to get the correct registry path for the current user
function Get-CurrentUserRegistryPath {
    param(
        [Parameter(Mandatory=$true)]
        [string]$RegistrySubPath
    )
    
    try {
        # Check if we're running with administrative privileges
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if ($isAdmin) {
            Write-LogMessage "INFO" "Running with administrative privileges - targeting actual user's registry hive" "TRUST"
            
            # Try multiple methods to get the current user
            $currentUser = $null
            $detectionMethod = ""
            
            # Method 1: Try quser command (if available)
            try {
                $quserPath = Get-Command "quser" -ErrorAction SilentlyContinue
                if ($quserPath) {
                    $activeSession = quser 2>$null | Where-Object { $_ -match '>' } | Select-Object -First 1
                    if ($activeSession) {
                        # Parse the quser output to get the username
                        $sessionData = $activeSession -replace '\s{2,}', ',' | ConvertFrom-Csv -Header "UserName","SessionName","ID","State","IdleTime","LogonTime"
                        $currentUser = $sessionData.UserName.Trim()
                        $detectionMethod = "quser"
                        Write-LogMessage "SUCCESS" "Detected active user session via quser: $currentUser" "TRUST"
                    }
                } else {
                    Write-LogMessage "INFO" "quser command not available in this environment" "TRUST"
                }
            } catch {
                Write-LogMessage "INFO" "quser command failed or not available: $_" "TRUST"
            }
            
            # Method 2: Try WMI query for logged-in users
            if (-not $currentUser) {
                try {
                    $loggedInUser = Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty UserName
                    if ($loggedInUser -and $loggedInUser.Contains('\')) {
                        $currentUser = $loggedInUser.Split('\')[1]
                        $detectionMethod = "WMI"
                        Write-LogMessage "SUCCESS" "Detected active user via WMI: $currentUser" "TRUST"
                    }
                } catch {
                    Write-LogMessage "INFO" "WMI user detection failed: $_" "TRUST"
                }
            }
            
            # Method 3: Try explorer.exe process owner
            if (-not $currentUser) {
                try {
                    $explorerProcess = Get-WmiObject -Class Win32_Process -Filter "Name='explorer.exe'" | Select-Object -First 1
                    if ($explorerProcess) {
                        $owner = $explorerProcess.GetOwner()
                        if ($owner.User) {
                            $currentUser = $owner.User
                            $detectionMethod = "ExplorerProcess"
                            Write-LogMessage "SUCCESS" "Detected active user via explorer.exe: $currentUser" "TRUST"
                        }
                    }
                } catch {
                    Write-LogMessage "INFO" "Explorer process user detection failed: $_" "TRUST"
                }
            }
            
            # Method 4: Fallback to environment variable
            if (-not $currentUser) {
                $currentUser = $env:USERNAME
                $detectionMethod = "Environment"
                Write-LogMessage "WARNING" "Using environment username as fallback: $currentUser" "TRUST"
            }
            
            Write-LogMessage "INFO" "User detection method used: $detectionMethod" "TRUST"
            
            # Retrieve the SID for the current user
            try {
                $userSID = (Get-LocalUser -Name $currentUser -ErrorAction Stop).SID.Value
                Write-LogMessage "SUCCESS" "Retrieved SID for user $currentUser`: $userSID" "TRUST"
                
                # Define the registry path under HKEY_USERS for the current user
                $registryPath = "Registry::HKEY_USERS\$userSID\$RegistrySubPath"
                Write-LogMessage "SUCCESS" "Using user-specific registry path: $registryPath" "TRUST"
                
                return @{
                    Success = $true
                    RegistryPath = $registryPath
                    UserName = $currentUser
                    UserSID = $userSID
                    Method = "HKEY_USERS"
                    DetectionMethod = $detectionMethod
                }
            } catch {
                Write-LogMessage "ERROR" "Failed to retrieve SID for user $currentUser`: $_" "TRUST"
                
                # Check if this might be a domain user
                if ($currentUser.Contains('\')) {
                    $domainUser = $currentUser.Split('\')[1]
                    try {
                        $userSID = (Get-LocalUser -Name $domainUser -ErrorAction Stop).SID.Value
                        $registryPath = "Registry::HKEY_USERS\$userSID\$RegistrySubPath"
                        Write-LogMessage "SUCCESS" "Retrieved SID for domain user $domainUser`: $userSID" "TRUST"
                        
                        return @{
                            Success = $true
                            RegistryPath = $registryPath
                            UserName = $domainUser
                            UserSID = $userSID
                            Method = "HKEY_USERS"
                            DetectionMethod = "$detectionMethod-Domain"
                        }
                    } catch {
                        Write-LogMessage "WARNING" "Failed to retrieve SID for domain user $domainUser`: $_" "TRUST"
                    }
                }
                
                # Fallback to HKCU (may not work correctly in admin context)
                $registryPath = "HKCU:\$RegistrySubPath"
                Write-LogMessage "WARNING" "Falling back to HKCU registry path (may not target correct user): $registryPath" "TRUST"
                
                return @{
                    Success = $false
                    RegistryPath = $registryPath
                    UserName = $currentUser
                    UserSID = $null
                    Method = "HKCU_FALLBACK"
                    DetectionMethod = $detectionMethod
                    Warning = "Could not retrieve user SID, using HKCU fallback"
                }
            }
        } else {
            # Not running as admin, HKCU should work correctly
            $registryPath = "HKCU:\$RegistrySubPath"
            $currentUser = $env:USERNAME
            Write-LogMessage "INFO" "Running in user context, using HKCU: $registryPath" "TRUST"
            
            return @{
                Success = $true
                RegistryPath = $registryPath
                UserName = $currentUser
                UserSID = $null
                Method = "HKCU"
                DetectionMethod = "UserContext"
            }
        }
    } catch {
        # Ultimate fallback
        $registryPath = "HKCU:\$RegistrySubPath"
        $currentUser = $env:USERNAME
        Write-LogMessage "ERROR" "Error in Get-CurrentUserRegistryPath, using fallback: $_" "TRUST"
        
        return @{
            Success = $false
            RegistryPath = $registryPath
            UserName = $currentUser
            UserSID = $null
            Method = "ERROR_FALLBACK"
            DetectionMethod = "Error"
            Error = $_.Exception.Message
        }
    }
}

# Function to add network share to trusted catalogs (current user only)
function Add-ToTrustedCatalogs {
    param(
        [Parameter(Mandatory=$true)]
        [string]$NetworkPath,
        
        [string]$RegistryPath = ""
    )
    
    try {
        Write-LogMessage "INFO" "Adding network share to trusted catalogs: $NetworkPath" "TRUST"
        Report-Progress -Status (Get-StatusCode "SETTING_TRUST") -Message "Configuring trusted catalogs"
        
        # Get the correct registry path for the current user
        if ($RegistryPath -eq "") {
            $baseRegistrySubPath = "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
            $registryResult = Get-CurrentUserRegistryPath -RegistrySubPath $baseRegistrySubPath
            $RegistryPath = $registryResult.RegistryPath
            $currentUser = $registryResult.UserName
            
            # Log the method used and any warnings
            Write-LogMessage "INFO" "Registry access method: $($registryResult.Method)" "TRUST"
            if ($registryResult.Warning) {
                Write-LogMessage "WARNING" $registryResult.Warning "TRUST"
            }
            if ($registryResult.Error) {
                Write-LogMessage "ERROR" "Registry path detection error: $($registryResult.Error)" "TRUST"
            }
        } else {
            $currentUser = $env:USERNAME
        }
        
        Write-LogMessage "INFO" "Target registry path: $RegistryPath" "TRUST"
        Write-LogMessage "INFO" "Target user: $currentUser" "TRUST"
        
        # Clear existing trusted catalogs (if any) to avoid conflicts
        try {
            if (Test-Path $RegistryPath) {
                Remove-Item -Path "$RegistryPath\*" -Force -ErrorAction SilentlyContinue
                Write-LogMessage "INFO" "Cleared existing trusted catalogs" "TRUST"
            }
        } catch {
            Write-LogMessage "WARNING" "Could not clear existing trusted catalogs: $_" "TRUST"
        }
        
        # Generate a new GUID for the catalog entry
        $catalogGuid = [guid]::NewGuid().ToString()
        $catalogPath = Join-Path $RegistryPath $catalogGuid
        
        # Ensure the base registry key exists
        if (-not (Test-Path $RegistryPath)) {
            try {
                New-Item -Path $RegistryPath -Force | Out-Null
                Write-LogMessage "SUCCESS" "Created registry path: $RegistryPath" "TRUST"
            } catch {
                $errorMsg = "Failed to create registry path: $RegistryPath - $_"
                Write-LogMessage "ERROR" $errorMsg "TRUST"
                Report-Error -ErrorCode (Get-ErrorCode "REGISTRY_ACCESS_FAILED") -ErrorMessage $errorMsg -Component "TRUST" -Exception $_
                return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "REGISTRY_ACCESS_FAILED")
            }
        }
        
        # Create a subkey for this catalog
        try {
            New-Item -Path $catalogPath -Force | Out-Null
            Write-LogMessage "SUCCESS" "Created catalog registry key: $catalogPath" "TRUST"
        } catch {
            $errorMsg = "Failed to create catalog registry key: $_"
            Write-LogMessage "ERROR" $errorMsg "TRUST"
            Report-Error -ErrorCode (Get-ErrorCode "REGISTRY_ACCESS_FAILED") -ErrorMessage $errorMsg -Component "TRUST" -Exception $_
            return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "REGISTRY_ACCESS_FAILED")
        }
        
        # Set required registry values
        # The "Flags" value of 3 indicates that the catalog is enabled and set to show in the add-in menu
        try {
            New-ItemProperty -Path $catalogPath -Name "Id"              -Value $catalogGuid -PropertyType String -Force | Out-Null
            New-ItemProperty -Path $catalogPath -Name "Url"             -Value $NetworkPath -PropertyType String -Force | Out-Null
            New-ItemProperty -Path $catalogPath -Name "Flags"           -Value 3           -PropertyType DWord  -Force | Out-Null
            New-ItemProperty -Path $catalogPath -Name "Type"            -Value 2           -PropertyType DWord  -Force | Out-Null
            New-ItemProperty -Path $catalogPath -Name "CatalogVersion"  -Value 2           -PropertyType DWord  -Force | Out-Null
            New-ItemProperty -Path $catalogPath -Name "SkipCatalogUpdate" -Value 0         -PropertyType DWord  -Force | Out-Null
            
            Write-LogMessage "SUCCESS" "Set trusted catalog registry values" "TRUST"
        } catch {
            $errorMsg = "Failed to set registry values: $_"
            Write-LogMessage "ERROR" $errorMsg "TRUST"
            Report-Error -ErrorCode (Get-ErrorCode "REGISTRY_ACCESS_FAILED") -ErrorMessage $errorMsg -Component "TRUST" -Exception $_
            return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "REGISTRY_ACCESS_FAILED")
        }
        
        # Verify the registry entries were created successfully
        try {
            $verifyPath = Test-Path $catalogPath
            $verifyId = Get-ItemProperty -Path $catalogPath -Name "Id" -ErrorAction SilentlyContinue
            $verifyUrl = Get-ItemProperty -Path $catalogPath -Name "Url" -ErrorAction SilentlyContinue
            
            if ($verifyPath -and $verifyId -and $verifyUrl) {
                Write-LogMessage "SUCCESS" "Verified trusted catalog registry entries" "TRUST"
            } else {
                Write-LogMessage "WARNING" "Could not verify all registry entries were created" "TRUST"
            }
        } catch {
            Write-LogMessage "WARNING" "Registry verification failed (non-critical): $_" "TRUST"
        }
        
        $result = @{
            NetworkPath = $NetworkPath
            CatalogGuid = $catalogGuid
            RegistryPath = $catalogPath
            CurrentUser = $currentUser
            RegistryMethod = $registryResult.Method
            UserSID = $registryResult.UserSID
        }
        
        Write-LogMessage "SUCCESS" "Added network share to trusted catalogs for user: $currentUser" "TRUST"
        Report-Progress -Status (Get-StatusCode "SETTING_TRUST") -Message "Trusted catalog configured" -PercentComplete 80 -AdditionalData $result
        
        return New-Result -Success $true -Message "Network share added to trusted catalogs" -Data $result
    }
    catch {
        $errorMsg = "Failed to add to trusted catalogs: $_"
        Write-LogMessage "ERROR" $errorMsg "TRUST"
        Report-Error -ErrorCode (Get-ErrorCode "REGISTRY_ACCESS_FAILED") -ErrorMessage $errorMsg -Component "TRUST" -Exception $_
        return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "REGISTRY_ACCESS_FAILED")
    }
}

# Function to check if share is in trusted locations
function Test-TrustedLocation {
    param(
        [Parameter(Mandatory=$true)]
        [string]$NetworkPath,
        
        [string]$RegistryPath = ""
    )
    
    try {
        Write-LogMessage "INFO" "Checking if network share is in trusted catalogs: $NetworkPath" "TRUST"
        
        # Get the correct registry path for the current user
        if ($RegistryPath -eq "") {
            $baseRegistrySubPath = "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
            $registryResult = Get-CurrentUserRegistryPath -RegistrySubPath $baseRegistrySubPath
            $RegistryPath = $registryResult.RegistryPath
            
            # Log the method used and any warnings
            Write-LogMessage "INFO" "Registry access method: $($registryResult.Method)" "TRUST"
            if ($registryResult.Warning) {
                Write-LogMessage "WARNING" $registryResult.Warning "TRUST"
            }
            if ($registryResult.Error) {
                Write-LogMessage "ERROR" "Registry path detection error: $($registryResult.Error)" "TRUST"
            }
        }
        
        Write-LogMessage "INFO" "Checking registry path: $RegistryPath" "TRUST"
        
        # Check if the registry key exists
        if (-not (Test-Path $RegistryPath)) {
            Write-LogMessage "INFO" "Trusted catalogs registry key doesn't exist" "TRUST"
            return @{
                IsTrusted = $false
                RegistryPath = $RegistryPath
                Issue = "Registry key not found"
            }
        }
        
        # Check if the share is already in trusted catalogs
        $found = $false
        $catalogInfo = $null
        
        try {
            Get-ChildItem -Path $RegistryPath -ErrorAction SilentlyContinue | ForEach-Object {
                $catalogPath = $_.PSPath
                $url = (Get-ItemProperty -Path $catalogPath -Name "Url" -ErrorAction SilentlyContinue).Url
                
                if ($url -eq $NetworkPath) {
                    $found = $true
                    $catalogInfo = @{
                        CatalogPath = $catalogPath
                        Url = $url
                        Id = (Get-ItemProperty -Path $catalogPath -Name "Id" -ErrorAction SilentlyContinue).Id
                        Flags = (Get-ItemProperty -Path $catalogPath -Name "Flags" -ErrorAction SilentlyContinue).Flags
                    }
                    Write-LogMessage "SUCCESS" "Network share is in trusted catalogs" "TRUST"
                }
            }
        } catch {
            Write-LogMessage "WARNING" "Error checking trusted catalogs: $_" "TRUST"
            return @{
                IsTrusted = $false
                RegistryPath = $RegistryPath
                Issue = "Error reading registry: $_"
            }
        }
        
        return @{
            IsTrusted = $found
            RegistryPath = $RegistryPath
            CatalogInfo = $catalogInfo
        }
    }
    catch {
        Write-LogMessage "ERROR" "Error checking trusted location: $_" "TRUST"
        return @{
            IsTrusted = $false
            RegistryPath = $RegistryPath
            Issue = "Exception: $_"
        }
    }
}

# Function to create share and trust configuration for all users (where possible)
function New-ShareAndTrustForAllUsers {
    param(
        [switch]$SkipExisting = $true,
        [switch]$CurrentUserOnly = $false
    )
    
    try {
        Write-LogMessage "INFO" "Creating shares and trust configuration for users" "SHARE"
        
        if ($CurrentUserOnly) {
            $currentUser = Get-CurrentUser
            if (-not $currentUser) {
                $errorMsg = "Could not determine current user"
                Write-LogMessage "ERROR" $errorMsg "SHARE"
                return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "USER_NOT_FOUND")
            }
            $users = @($currentUser)
        } else {
            $users = Get-SystemUsers -ExcludeSystemAccounts
        }
        
        $results = @()
        
        foreach ($user in $users) {
            Write-LogMessage "INFO" "Processing share and trust for user: $($user.Username)" "SHARE"
            
            $userResult = @{
                Username = $user.Username
                ShareResult = $null
                TrustResult = $null
                OverallSuccess = $false
            }
            
            # Create network share for this user
            $shareResult = New-UserNetworkShare -UserInfo $user
            $userResult.ShareResult = $shareResult
            
            if ($shareResult.Success) {
                # Add to trusted catalogs (only works for current user)
                if ($user.Username -eq $env:USERNAME) {
                    $trustResult = Add-ToTrustedCatalogs -NetworkPath $shareResult.Data.NetworkPath
                    $userResult.TrustResult = $trustResult
                    $userResult.OverallSuccess = $trustResult.Success
                } else {
                    Write-LogMessage "WARNING" "Cannot configure trust center for other users, skipping trust configuration for: $($user.Username)" "SHARE"
                    $userResult.TrustResult = New-Result -Success $false -Message "Cannot configure trust for other users"
                    $userResult.OverallSuccess = $shareResult.Success # Share creation is enough for other users
                }
            } else {
                $userResult.OverallSuccess = $false
            }
            
            $results += $userResult
            
            if ($userResult.OverallSuccess) {
                Write-LogMessage "SUCCESS" "Completed share and trust setup for user: $($user.Username)" "SHARE"
            } else {
                Write-LogMessage "ERROR" "Failed share and trust setup for user: $($user.Username)" "SHARE"
            }
        }
        
        $successCount = ($results | Where-Object { $_.OverallSuccess }).Count
        $totalCount = $results.Count
        
        Write-LogMessage "SUCCESS" "Share and trust setup completed: $successCount/$totalCount users" "SHARE"
        
        return New-Result -Success ($successCount -gt 0) -Message "Share and trust setup completed for $successCount/$totalCount users" -Data @{
            Results = $results
            SuccessCount = $successCount
            TotalCount = $totalCount
        }
    }
    catch {
        $errorMsg = "Failed to create shares and trust for users: $_"
        Write-LogMessage "ERROR" $errorMsg "SHARE"
        return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "NETWORK_SHARE_FAILED")
    }
}

# Note: Export-ModuleMember is not needed for .ps1 files that are dot-sourced 