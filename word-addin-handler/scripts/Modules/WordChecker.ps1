# Office Add-in Setup - Word Installation Checker Module
# This module handles verification of Word installation and compatibility

# Import required modules
. "$PSScriptRoot\..\Core\Constants.ps1"
. "$PSScriptRoot\..\Core\Logging.ps1"

# Function to check if Word is installed on the device
function Test-WordInstalled {
    param(
        [switch]$CheckAllVersions = $false
    )
    
    try {
        Write-LogMessage "INFO" "Checking Microsoft Word installation" "WORD"
        Report-Progress -Status (Get-StatusCode "CHECKING_WORD") -Message "Verifying Word installation"
        
        $wordInstalled = $false
        $wordVersion = $null
        $wordPath = $null
        $installationMethod = $null
        
        # Method 1: Check Word application registry keys for Click-to-Run
        $clickToRunPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        if (Test-Path $clickToRunPath) {
            try {
                $platform = (Get-ItemProperty $clickToRunPath -Name "Platform" -ErrorAction SilentlyContinue).Platform
                $versionToReport = (Get-ItemProperty $clickToRunPath -Name "VersionToReport" -ErrorAction SilentlyContinue).VersionToReport
                
                if ($platform -and $versionToReport) {
                    Write-LogMessage "SUCCESS" "Found Click-to-Run Office installation: $versionToReport ($platform)" "WORD"
                    $wordInstalled = $true
                    $wordVersion = $versionToReport
                    $installationMethod = "Click-to-Run"
                }
            } catch {
                Write-LogMessage "DEBUG" "Could not read Click-to-Run registry: $_" "WORD"
            }
        }
        
        # Method 2: Check traditional MSI installation
        if (-not $wordInstalled) {
            $msiPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot",
                "HKLM:\SOFTWARE\Microsoft\Office\15.0\Word\InstallRoot",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Word\InstallRoot",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Word\InstallRoot"
            )
            
            foreach ($path in $msiPaths) {
                if (Test-Path $path) {
                    try {
                        $installPath = (Get-ItemProperty $path -Name "Path" -ErrorAction SilentlyContinue).Path
                        if ($installPath -and (Test-Path (Join-Path $installPath "WINWORD.EXE"))) {
                            Write-LogMessage "SUCCESS" "Found MSI Office installation at: $installPath" "WORD"
                            $wordInstalled = $true
                            $wordPath = Join-Path $installPath "WINWORD.EXE"
                            $installationMethod = "MSI"
                            $wordVersion = if ($path -match "16\.0") { "2016+" } elseif ($path -match "15\.0") { "2013" } else { "Unknown" }
                            break
                        }
                    } catch {
                        Write-LogMessage "DEBUG" "Could not read MSI registry path $path : $_" "WORD"
                    }
                }
            }
        }
        
        # Method 3: Check common installation paths
        if (-not $wordInstalled) {
            $commonPaths = @(
                "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
                "C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
                "C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
                "C:\Program Files\Microsoft Office\Office15\WINWORD.EXE",
                "C:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE"
            )
            
            foreach ($path in $commonPaths) {
                if (Test-Path $path) {
                    Write-LogMessage "SUCCESS" "Found Word executable at: $path" "WORD"
                    $wordInstalled = $true
                    $wordPath = $path
                    $installationMethod = "FileSystem"
                    $wordVersion = if ($path -match "Office16") { "2016+" } elseif ($path -match "Office15") { "2013" } else { "Unknown" }
                    break
                }
            }
        }
        
        # Method 4: Try to create COM object to verify functionality
        if ($wordInstalled) {
            try {
                Write-LogMessage "INFO" "Testing Word COM object functionality" "WORD"
                $word = New-Object -ComObject Word.Application -ErrorAction Stop
                $word.Visible = $false
                $comVersion = $word.Version
                $word.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
                
                Write-LogMessage "SUCCESS" "Word COM object test successful. Version: $comVersion" "WORD"
                $wordVersion = $comVersion
            } catch {
                Write-LogMessage "WARNING" "Word is installed but COM object test failed: $_" "WORD"
                # Don't fail here, Word might still work for our purposes
            }
        }
        
        # Create result object
        $result = @{
            IsInstalled = $wordInstalled
            Version = $wordVersion
            Path = $wordPath
            InstallationMethod = $installationMethod
            IsCompatible = $wordInstalled # For now, assume any Word installation is compatible
        }
        
        if ($wordInstalled) {
            Write-LogMessage "SUCCESS" "Microsoft Word is installed and functional" "WORD"
            Report-Progress -Status (Get-StatusCode "CHECKING_WORD") -Message "Word installation verified" -PercentComplete 20 -AdditionalData $result
            return New-Result -Success $true -Message "Word is installed and functional" -Data $result
        } else {
            $errorMsg = "Microsoft Word is not installed on this device"
            Write-LogMessage "ERROR" $errorMsg "WORD"
            Report-Error -ErrorCode (Get-ErrorCode "WORD_NOT_INSTALLED") -ErrorMessage $errorMsg -Component "WORD"
            return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "WORD_NOT_INSTALLED")
        }
    }
    catch {
        $errorMsg = "Failed to check Word installation: $_"
        Write-LogMessage "ERROR" $errorMsg "WORD"
        Report-Error -ErrorCode (Get-ErrorCode "WORD_NOT_INSTALLED") -ErrorMessage $errorMsg -Component "WORD" -Exception $_
        return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "WORD_NOT_INSTALLED")
    }
}

# Function to get Word version information
function Get-WordVersion {
    try {
        Write-LogMessage "INFO" "Getting detailed Word version information" "WORD"
        
        $versionInfo = @{
            Version = $null
            Build = $null
            Architecture = $null
            Channel = $null
            LastUpdate = $null
        }
        
        # Try to get version from COM object
        try {
            $word = New-Object -ComObject Word.Application -ErrorAction Stop
            $word.Visible = $false
            $versionInfo.Version = $word.Version
            $versionInfo.Build = $word.Build
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        } catch {
            Write-LogMessage "WARNING" "Could not get version from COM object: $_" "WORD"
        }
        
        # Try to get additional info from Click-to-Run registry
        $clickToRunPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        if (Test-Path $clickToRunPath) {
            try {
                $versionInfo.Architecture = (Get-ItemProperty $clickToRunPath -Name "Platform" -ErrorAction SilentlyContinue).Platform
                $versionInfo.Channel = (Get-ItemProperty $clickToRunPath -Name "CDNBaseUrl" -ErrorAction SilentlyContinue).CDNBaseUrl
            } catch {
                Write-LogMessage "DEBUG" "Could not get Click-to-Run details: $_" "WORD"
            }
        }
        
        Write-LogMessage "SUCCESS" "Word version info: $($versionInfo.Version) Build: $($versionInfo.Build)" "WORD"
        return $versionInfo
    }
    catch {
        Write-LogMessage "ERROR" "Failed to get Word version: $_" "WORD"
        return $null
    }
}

# Function to check if Word supports the required add-in features
function Test-WordAddinCompatibility {
    param(
        [string]$RequiredVersion = "16.0"
    )
    
    try {
        Write-LogMessage "INFO" "Checking Word add-in compatibility" "WORD"
        
        $wordCheck = Test-WordInstalled
        if (-not $wordCheck.Success) {
            return $wordCheck
        }
        
        $versionInfo = Get-WordVersion
        if (-not $versionInfo -or -not $versionInfo.Version) {
            $errorMsg = "Could not determine Word version for compatibility check"
            Write-LogMessage "ERROR" $errorMsg "WORD"
            return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "WORD_NOT_INSTALLED")
        }
        
        # Check if version meets minimum requirements
        $currentVersion = [Version]$versionInfo.Version
        $minVersion = [Version]$RequiredVersion
        
        if ($currentVersion -ge $minVersion) {
            Write-LogMessage "SUCCESS" "Word version $($versionInfo.Version) meets minimum requirements ($RequiredVersion)" "WORD"
            return New-Result -Success $true -Message "Word is compatible with add-ins" -Data $versionInfo
        } else {
            $errorMsg = "Word version $($versionInfo.Version) does not meet minimum requirements ($RequiredVersion)"
            Write-LogMessage "ERROR" $errorMsg "WORD"
            return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "WORD_NOT_INSTALLED")
        }
    }
    catch {
        $errorMsg = "Failed to check Word compatibility: $_"
        Write-LogMessage "ERROR" $errorMsg "WORD"
        return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "WORD_NOT_INSTALLED")
    }
}

# Note: Export-ModuleMember is not needed for .ps1 files that are dot-sourced 