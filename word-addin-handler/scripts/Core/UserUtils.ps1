# Office Add-in Setup - User Context Management Module
# This module handles user detection and multi-user scenarios

# Import required modules
. "$PSScriptRoot\Constants.ps1"
. "$PSScriptRoot\Logging.ps1"

# Function to get all users on the system
function Get-SystemUsers {
    param(
        [switch]$ExcludeSystemAccounts = $true
    )
    
    try {
        Write-LogMessage "INFO" "Discovering users on the system" "USER"
        
        $users = @()
        $excludedUsers = @("Administrator", "Public", "Default", "DefaultAppPool", "All Users")
        
        # Get users from registry
        $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        if (Test-Path $profileListPath) {
            Get-ChildItem $profileListPath | ForEach-Object {
                $profilePath = (Get-ItemProperty $_.PSPath -Name "ProfileImagePath" -ErrorAction SilentlyContinue).ProfileImagePath
                if ($profilePath -and (Test-Path $profilePath)) {
                    $username = Split-Path $profilePath -Leaf
                    if (-not $ExcludeSystemAccounts -or $username -notin $excludedUsers) {
                        $users += @{
                            Username = $username
                            ProfilePath = $profilePath
                            DocumentsPath = Join-Path $profilePath "Documents"
                            SID = Split-Path $_.PSPath -Leaf
                            IsActive = (Test-Path (Join-Path $profilePath "NTUSER.DAT"))
                        }
                    }
                }
            }
        }
        
        # Also check C:\Users directory
        if (Test-Path "C:\Users") {
            Get-ChildItem "C:\Users" -Directory | ForEach-Object {
                $username = $_.Name
                if (-not $ExcludeSystemAccounts -or $username -notin $excludedUsers) {
                    # Check if this user is already in our list
                    $existingUser = $users | Where-Object { $_.Username -eq $username }
                    if (-not $existingUser -and (Test-Path (Join-Path $_.FullName "NTUSER.DAT"))) {
                        $users += @{
                            Username = $username
                            ProfilePath = $_.FullName
                            DocumentsPath = Join-Path $_.FullName "Documents"
                            SID = $null
                            IsActive = $true
                        }
                    }
                }
            }
        }
        
        Write-LogMessage "SUCCESS" "Found $($users.Count) users on the system" "USER"
        return $users
    }
    catch {
        Write-LogMessage "ERROR" "Failed to discover system users: $_" "USER"
        return @()
    }
}

# Function to get current active user
function Get-CurrentUser {
    param(
        [switch]$PreferLoggedOnUser = $true
    )
    
    try {
        Write-LogMessage "INFO" "Determining current user context" "USER"
        
        $currentUser = $null
        
        # Method 1: Try to get the user who owns explorer.exe process
        if ($PreferLoggedOnUser) {
            try {
                $explorerProcess = Get-WmiObject -Class Win32_Process -Filter "Name='explorer.exe'" | 
                    Where-Object { 
                        $owner = $_.GetOwner()
                        $owner.User -and $owner.User -ne "Administrator" 
                    } | 
                    Select-Object -First 1
                
                if ($explorerProcess) {
                    $owner = $explorerProcess.GetOwner()
                    $username = $owner.User
                    $domain = $owner.Domain
                    
                    Write-LogMessage "SUCCESS" "Detected logged-on user from explorer.exe: $domain\$username" "USER"
                    
                    $userProfile = "C:\Users\$username"
                    if (Test-Path $userProfile) {
                        $currentUser = @{
                            Username = $username
                            Domain = $domain
                            ProfilePath = $userProfile
                            DocumentsPath = Join-Path $userProfile "Documents"
                            DetectionMethod = "ExplorerProcess"
                            IsLoggedOn = $true
                        }
                    }
                }
            } catch {
                Write-LogMessage "WARNING" "Could not get user from explorer.exe process: $_" "USER"
            }
        }
        
        # Method 2: Check who's logged on console
        if (-not $currentUser) {
            try {
                $loggedOnUsers = quser 2>$null | Where-Object { $_ -match "console" -and $_ -notmatch "Administrator" }
                if ($loggedOnUsers) {
                    $userLine = $loggedOnUsers | Select-Object -First 1
                    $parts = $userLine -split '\s+'
                    if ($parts.Count -ge 2) {
                        $username = $parts[1]
                        
                        Write-LogMessage "SUCCESS" "Detected console user: $username" "USER"
                        
                        $userProfile = "C:\Users\$username"
                        if (Test-Path $userProfile) {
                            $currentUser = @{
                                Username = $username
                                Domain = $env:USERDOMAIN
                                ProfilePath = $userProfile
                                DocumentsPath = Join-Path $userProfile "Documents"
                                DetectionMethod = "ConsoleSession"
                                IsLoggedOn = $true
                            }
                        }
                    }
                }
            } catch {
                Write-LogMessage "WARNING" "Could not get console user: $_" "USER"
            }
        }
        
        # Method 3: Use environment variables (may be admin when running elevated)
        if (-not $currentUser) {
            $username = $env:USERNAME
            $userProfile = "C:\Users\$username"
            
            if (Test-Path $userProfile) {
                Write-LogMessage "WARNING" "Using environment username (may be admin): $username" "USER"
                $currentUser = @{
                    Username = $username
                    Domain = $env:USERDOMAIN
                    ProfilePath = $userProfile
                    DocumentsPath = Join-Path $userProfile "Documents"
                    DetectionMethod = "Environment"
                    IsLoggedOn = $false
                }
            }
        }
        
        # Method 4: Find most recently accessed user profile
        if (-not $currentUser) {
            try {
                $recentProfile = Get-ChildItem "C:\Users" -Directory | 
                    Where-Object { 
                        $_.Name -notin @("Administrator", "Public", "Default", "DefaultAppPool") -and
                        (Test-Path (Join-Path $_.FullName "NTUSER.DAT"))
                    } | 
                    Sort-Object LastWriteTime -Descending | 
                    Select-Object -First 1
                
                if ($recentProfile) {
                    Write-LogMessage "WARNING" "Using most recent user profile: $($recentProfile.Name)" "USER"
                    $currentUser = @{
                        Username = $recentProfile.Name
                        Domain = $env:USERDOMAIN
                        ProfilePath = $recentProfile.FullName
                        DocumentsPath = Join-Path $recentProfile.FullName "Documents"
                        DetectionMethod = "RecentProfile"
                        IsLoggedOn = $false
                    }
                }
            } catch {
                Write-LogMessage "WARNING" "Could not detect user from recent profiles: $_" "USER"
            }
        }
        
        if ($currentUser) {
            Write-LogMessage "SUCCESS" "Current user determined: $($currentUser.Username) (Method: $($currentUser.DetectionMethod))" "USER"
            return $currentUser
        } else {
            Write-LogMessage "ERROR" "Could not determine current user" "USER"
            return $null
        }
    }
    catch {
        Write-LogMessage "ERROR" "Failed to determine current user: $_" "USER"
        return $null
    }
}

# Function to test if current process has admin rights
function Test-AdminRights {
    try {
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
        Write-LogMessage "INFO" "Admin rights check: $isAdmin" "USER"
        return $isAdmin
    }
    catch {
        Write-LogMessage "WARNING" "Could not check admin rights: $_" "USER"
        return $false
    }
}

# Function to test user access to registry paths
function Test-UserRegistryAccess {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Username,
        
        [string]$RegistryPath = "HKCU:\Software\Microsoft\Office"
    )
    
    try {
        # For current user, test directly
        if ($Username -eq $env:USERNAME) {
            $testPath = $RegistryPath
            if (Test-Path $testPath) {
                Write-LogMessage "SUCCESS" "Current user has registry access to $testPath" "USER"
                return $true
            } else {
                # Try to create the path
                try {
                    New-Item -Path $testPath -Force -ErrorAction Stop | Out-Null
                    Write-LogMessage "SUCCESS" "Created registry path for current user: $testPath" "USER"
                    return $true
                } catch {
                    Write-LogMessage "ERROR" "Cannot access or create registry path: $testPath" "USER"
                    return $false
                }
            }
        } else {
            # For other users, we would need to load their hive
            # This is complex and typically requires admin rights
            Write-LogMessage "WARNING" "Cannot test registry access for other users without loading their hive" "USER"
            return $false
        }
    }
    catch {
        Write-LogMessage "ERROR" "Error testing registry access for user $Username : $_" "USER"
        return $false
    }
}

# Function to test user access to file system paths
function Test-UserFileSystemAccess {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$UserInfo,
        
        [string]$TestPath = ""
    )
    
    try {
        if ($TestPath -eq "") {
            $TestPath = $UserInfo.DocumentsPath
        }
        
        # Test if we can read the path
        if (-not (Test-Path $TestPath)) {
            # Try to create the directory
            try {
                New-Item -Path $TestPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Write-LogMessage "SUCCESS" "Created directory for user: $TestPath" "USER"
            } catch {
                Write-LogMessage "ERROR" "Cannot create directory: $TestPath - $_" "USER"
                return $false
            }
        }
        
        # Test write access
        $testFile = Join-Path $TestPath "test-access-$(Get-Random).tmp"
        try {
            "test" | Out-File -FilePath $testFile -ErrorAction Stop
            Remove-Item -Path $testFile -ErrorAction SilentlyContinue
            Write-LogMessage "SUCCESS" "User has read/write access to: $TestPath" "USER"
            return $true
        } catch {
            Write-LogMessage "ERROR" "No write access to: $TestPath - $_" "USER"
            return $false
        }
    }
    catch {
        Write-LogMessage "ERROR" "Error testing file system access: $_" "USER"
        return $false
    }
}

# Note: Export-ModuleMember is not needed for .ps1 files that are dot-sourced 