param(
    [Parameter(Mandatory=$true)]
    [string]$documentName,
    
    [Parameter(Mandatory=$false)]
    [string]$documentUrl = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$AllUsers = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipDocumentOpen = $false,
    
    [Parameter(Mandatory=$false)]
    [string]$LogPath = "",
    
    [Parameter(Mandatory=$false)]
    [string]$ApiEndpoint = ""
)

# Office Add-in Setup - Main Orchestration Script
# This script coordinates all modules to set up Office add-ins with improved error handling

# Set script directory and import core modules
$ScriptRoot = $PSScriptRoot
$CorePath = Join-Path $ScriptRoot "Core"
$ModulesPath = Join-Path $ScriptRoot "Modules"

# Import core modules
. "$CorePath\Constants.ps1"
. "$CorePath\Logging.ps1"
. "$CorePath\UserUtils.ps1"

# Import feature modules
. "$ModulesPath\WordChecker.ps1"
. "$ModulesPath\ManifestManager.ps1"
. "$ModulesPath\ShareManager.ps1"

# Global variables for tracking
$Global:SetupResults = @{
    StartTime = Get-Date
    WordCheck = $null
    ManifestCreation = $null
    ShareConfiguration = $null
    TrustConfiguration = $null
    DocumentOpen = $null
    OverallSuccess = $false
    Errors = @()
}

# Function to send progress updates to API if endpoint is provided
function Send-ApiUpdate {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$UpdateData
    )
    
    if ($ApiEndpoint -eq "") {
        return
    }
    
    try {
        $jsonData = $UpdateData | ConvertTo-Json -Depth 10
        Invoke-RestMethod -Uri $ApiEndpoint -Method POST -Body $jsonData -ContentType "application/json" -TimeoutSec 10
        Write-LogMessage "DEBUG" "API update sent successfully" "API"
    } catch {
        Write-LogMessage "WARNING" "Failed to send API update: $_" "API"
    }
}

# Function to initialize the setup process
function Initialize-Setup {
    try {
        Write-Host "=== Office Add-in Setup - Modular Version ===" -ForegroundColor Cyan
        Write-Host "Document: $documentName" -ForegroundColor Green
        if ($documentUrl -ne "") {
            Write-Host "URL: $documentUrl" -ForegroundColor Green
        }
        Write-Host "All Users: $AllUsers" -ForegroundColor Green
        Write-Host "Skip Document Open: $SkipDocumentOpen" -ForegroundColor Green
        
        # Initialize logging with optional API callback
        $progressCallback = if ($ApiEndpoint -ne "") {
            { param($data) Send-ApiUpdate -UpdateData $data }
        } else {
            $null
        }
        
        $logFile = Initialize-Logging -LogPath $LogPath -ProgressCallback $progressCallback
        Write-LogMessage "INFO" "Setup initialized. Log file: $logFile" "MAIN"
        
        Report-Progress -Status (Get-StatusCode "STARTING") -Message "Office Add-in setup starting" -PercentComplete 0 -AdditionalData @{
            DocumentName = $documentName
            DocumentUrl = $documentUrl
            AllUsers = $AllUsers
            SkipDocumentOpen = $SkipDocumentOpen
            LogFile = $logFile
        }
        
        return $true
    } catch {
        Write-Error "Failed to initialize setup: $_"
        return $false
    }
}

# Function to check prerequisites
function Test-Prerequisites {
    try {
        Write-LogMessage "INFO" "Checking prerequisites" "MAIN"
        
        # Check if running with admin rights
        $isAdmin = Test-AdminRights
        if (-not $isAdmin) {
            $errorMsg = "This script requires administrator privileges"
            Write-LogMessage "ERROR" $errorMsg "MAIN"
            Report-Error -ErrorCode (Get-ErrorCode "INSUFFICIENT_PERMISSIONS") -ErrorMessage $errorMsg -Component "MAIN"
            return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "INSUFFICIENT_PERMISSIONS")
        }
        
        # Check Word installation
        Write-LogMessage "INFO" "Checking Word installation" "MAIN"
        $wordCheck = Test-WordInstalled
        $Global:SetupResults.WordCheck = $wordCheck
        
        if (-not $wordCheck.Success) {
            $Global:SetupResults.Errors += $wordCheck
            return $wordCheck
        }
        
        # Get current user information
        $currentUser = Get-CurrentUser
        if (-not $currentUser) {
            $errorMsg = "Could not determine current user context"
            Write-LogMessage "ERROR" $errorMsg "MAIN"
            Report-Error -ErrorCode (Get-ErrorCode "USER_NOT_FOUND") -ErrorMessage $errorMsg -Component "MAIN"
            return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "USER_NOT_FOUND")
        }
        
        Write-LogMessage "SUCCESS" "Prerequisites check completed" "MAIN"
        return New-Result -Success $true -Message "Prerequisites satisfied" -Data @{
            WordCheck = $wordCheck
            CurrentUser = $currentUser
            IsAdmin = $isAdmin
        }
    } catch {
        $errorMsg = "Failed during prerequisites check: $_"
        Write-LogMessage "ERROR" $errorMsg "MAIN"
        Report-Error -ErrorCode (Get-ErrorCode "ADDIN_CONFIGURATION_FAILED") -ErrorMessage $errorMsg -Component "MAIN" -Exception $_
        return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "ADDIN_CONFIGURATION_FAILED")
    }
}

# Function to set up manifests
function Setup-Manifests {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$CurrentUser
    )
    
    try {
        Write-LogMessage "INFO" "Setting up manifest files" "MAIN"
        
        if ($AllUsers) {
            Write-LogMessage "INFO" "Creating manifests for all users" "MAIN"
            $manifestResult = New-ManifestForAllUsers -SkipExisting
        } else {
            Write-LogMessage "INFO" "Creating manifest for current user only" "MAIN"
            $manifestResult = New-ManifestFile -UserInfo $CurrentUser
        }
        
        $Global:SetupResults.ManifestCreation = $manifestResult
        
        if (-not $manifestResult.Success) {
            $Global:SetupResults.Errors += $manifestResult
            return $manifestResult
        }
        
        Write-LogMessage "SUCCESS" "Manifest setup completed" "MAIN"
        return $manifestResult
    } catch {
        $errorMsg = "Failed during manifest setup: $_"
        Write-LogMessage "ERROR" $errorMsg "MAIN"
        Report-Error -ErrorCode (Get-ErrorCode "MANIFEST_CREATION_FAILED") -ErrorMessage $errorMsg -Component "MAIN" -Exception $_
        return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "MANIFEST_CREATION_FAILED")
    }
}

# Function to set up network shares and trust configuration
function Setup-SharesAndTrust {
    try {
        Write-LogMessage "INFO" "Setting up network shares and trust configuration" "MAIN"
        
        if ($AllUsers) {
            Write-LogMessage "INFO" "Creating shares and trust for all users" "MAIN"
            $shareResult = New-ShareAndTrustForAllUsers -SkipExisting
        } else {
            Write-LogMessage "INFO" "Creating shares and trust for current user only" "MAIN"
            $shareResult = New-ShareAndTrustForAllUsers -CurrentUserOnly -SkipExisting
        }
        
        $Global:SetupResults.ShareConfiguration = $shareResult
        
        if (-not $shareResult.Success) {
            $Global:SetupResults.Errors += $shareResult
            return $shareResult
        }
        
        Write-LogMessage "SUCCESS" "Share and trust setup completed" "MAIN"
        return $shareResult
    } catch {
        $errorMsg = "Failed during share and trust setup: $_"
        Write-LogMessage "ERROR" $errorMsg "MAIN"
        Report-Error -ErrorCode (Get-ErrorCode "NETWORK_SHARE_FAILED") -ErrorMessage $errorMsg -Component "MAIN" -Exception $_
        return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "NETWORK_SHARE_FAILED")
    }
}

# Function to open document (placeholder for future Word automation module)
function Open-Document {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DocumentName,
        
        [string]$DocumentUrl = ""
    )
    
    try {
        if ($SkipDocumentOpen) {
            Write-LogMessage "INFO" "Skipping document open as requested" "MAIN"
            return New-Result -Success $true -Message "Document open skipped" -Data @{ Skipped = $true }
        }
        
        Write-LogMessage "INFO" "Opening document: $DocumentName" "MAIN"
        Report-Progress -Status (Get-StatusCode "OPENING_DOCUMENT") -Message "Opening Word document"
        
        # For now, this is a placeholder
        # The full Word automation logic would go here
        Write-LogMessage "INFO" "Document opening functionality to be implemented" "MAIN"
        
        $result = New-Result -Success $true -Message "Document handling completed" -Data @{
            DocumentName = $DocumentName
            DocumentUrl = $DocumentUrl
            Note = "Full automation to be implemented"
        }
        
        $Global:SetupResults.DocumentOpen = $result
        return $result
    } catch {
        $errorMsg = "Failed to open document: $_"
        Write-LogMessage "ERROR" $errorMsg "MAIN"
        Report-Error -ErrorCode (Get-ErrorCode "OPEN_DOCUMENT_FAILED") -ErrorMessage $errorMsg -Component "MAIN" -Exception $_
        return New-Result -Success $false -Message $errorMsg -ErrorCode (Get-ErrorCode "OPEN_DOCUMENT_FAILED")
    }
}

# Function to generate final setup summary
function New-SetupSummary {
    try {
        $endTime = Get-Date
        $duration = $endTime - $Global:SetupResults.StartTime
        
        $summary = @{
            StartTime = $Global:SetupResults.StartTime
            EndTime = $endTime
            Duration = $duration
            OverallSuccess = $Global:SetupResults.OverallSuccess
            Results = @{
                WordCheck = $Global:SetupResults.WordCheck
                ManifestCreation = $Global:SetupResults.ManifestCreation
                ShareConfiguration = $Global:SetupResults.ShareConfiguration
                TrustConfiguration = $Global:SetupResults.TrustConfiguration
                DocumentOpen = $Global:SetupResults.DocumentOpen
            }
            Errors = $Global:SetupResults.Errors
            Parameters = @{
                DocumentName = $documentName
                DocumentUrl = $documentUrl
                AllUsers = $AllUsers
                SkipDocumentOpen = $SkipDocumentOpen
            }
        }
        
        Write-LogMessage "INFO" "Setup Summary:" "MAIN"
        Write-LogMessage "INFO" "- Duration: $($duration.TotalSeconds) seconds" "MAIN"
        Write-LogMessage "INFO" "- Overall Success: $($summary.OverallSuccess)" "MAIN"
        Write-LogMessage "INFO" "- Errors Count: $($summary.Errors.Count)" "MAIN"
        
        # Send final status update
        $finalStatus = if ($summary.OverallSuccess) { Get-StatusCode "COMPLETED" } else { Get-StatusCode "FAILED" }
        $finalMessage = if ($summary.OverallSuccess) { "Setup completed successfully" } else { "Setup failed with errors" }
        
        Report-Progress -Status $finalStatus -Message $finalMessage -PercentComplete 100 -AdditionalData $summary
        
        return $summary
    } catch {
        Write-LogMessage "ERROR" "Failed to generate setup summary: $_" "MAIN"
        return $null
    }
}

# Main execution function
function Start-OfficeAddinSetup {
    try {
        # Initialize
        if (-not (Initialize-Setup)) {
            exit 1
        }
        
        # Check prerequisites
        Write-LogMessage "INFO" "Starting prerequisites check" "MAIN"
        $prereqResult = Test-Prerequisites
        if (-not $prereqResult.Success) {
            Write-LogMessage "ERROR" "Prerequisites check failed: $($prereqResult.Message)" "MAIN"
            $Global:SetupResults.OverallSuccess = $false
            exit $prereqResult.ErrorCode
        }
        
        $currentUser = $prereqResult.Data.CurrentUser
        Write-LogMessage "SUCCESS" "Prerequisites satisfied for user: $($currentUser.Username)" "MAIN"
        
        # Setup manifests
        Write-LogMessage "INFO" "Starting manifest setup" "MAIN"
        $manifestResult = Setup-Manifests -CurrentUser $currentUser
        if (-not $manifestResult.Success) {
            Write-LogMessage "ERROR" "Manifest setup failed: $($manifestResult.Message)" "MAIN"
            $Global:SetupResults.OverallSuccess = $false
            exit $manifestResult.ErrorCode
        }
        
        # Setup shares and trust
        Write-LogMessage "INFO" "Starting share and trust setup" "MAIN"
        $shareResult = Setup-SharesAndTrust
        if (-not $shareResult.Success) {
            Write-LogMessage "ERROR" "Share and trust setup failed: $($shareResult.Message)" "MAIN"
            $Global:SetupResults.OverallSuccess = $false
            exit $shareResult.ErrorCode
        }
        
        # Open document
        Write-LogMessage "INFO" "Starting document handling" "MAIN"
        $documentResult = Open-Document -DocumentName $documentName -DocumentUrl $documentUrl
        if (-not $documentResult.Success) {
            Write-LogMessage "WARNING" "Document handling failed: $($documentResult.Message)" "MAIN"
            # Don't exit on document failure, as the main setup is complete
        }
        
        # Overall success
        $Global:SetupResults.OverallSuccess = $true
        Write-LogMessage "SUCCESS" "Office Add-in setup completed successfully!" "MAIN"
        
        # Generate and display summary
        $summary = New-SetupSummary
        
        # Display user instructions
        Write-Host "`n=== Setup Completed Successfully ===" -ForegroundColor Green
        Write-Host "Configuration details:" -ForegroundColor Cyan
        Write-Host "- Target User: $($currentUser.Username)" -ForegroundColor White
        
        if ($AllUsers) {
            Write-Host "- Configured for: All system users" -ForegroundColor White
        } else {
            Write-Host "- Configured for: Current user only" -ForegroundColor White
        }
        
        Write-Host "`nNext Steps:" -ForegroundColor Yellow
        Write-Host "1. Open Microsoft Word" -ForegroundColor White
        Write-Host "2. Go to File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs" -ForegroundColor White
        Write-Host "3. Verify the network share is listed and 'Show in Menu' is checked" -ForegroundColor White
        Write-Host "4. Click OK and restart Word" -ForegroundColor White
        Write-Host "5. The add-in should appear in the ribbon or Insert > My Add-ins" -ForegroundColor White
        
        exit 0
        
    } catch {
        $errorMsg = "Critical error in setup process: $_"
        Write-LogMessage "ERROR" $errorMsg "MAIN"
        Report-Error -ErrorCode (Get-ErrorCode "ADDIN_CONFIGURATION_FAILED") -ErrorMessage $errorMsg -Component "MAIN" -Exception $_
        
        $Global:SetupResults.OverallSuccess = $false
        New-SetupSummary
        
        Write-Host "`n=== Setup Failed ===" -ForegroundColor Red
        Write-Host "Error: $errorMsg" -ForegroundColor Red
        Write-Host "Check the log file for detailed information." -ForegroundColor Yellow
        
        exit (Get-ErrorCode "ADDIN_CONFIGURATION_FAILED")
    }
}

# Script entry point
try {
    # Check if running with required privileges
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "This script requires Administrator privileges!"
        Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
        exit (Get-ErrorCode "INSUFFICIENT_PERMISSIONS")
    }
    
    # Start the setup process
    Start-OfficeAddinSetup
    
} catch {
    Write-Error "Fatal error: $_"
    Write-Host "Setup failed with a critical error. Please check the logs." -ForegroundColor Red
    exit (Get-ErrorCode "ADDIN_CONFIGURATION_FAILED")
} 