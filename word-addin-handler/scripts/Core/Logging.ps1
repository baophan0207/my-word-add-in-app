# Office Add-in Setup - Logging and Progress Reporting Module
# This module handles all logging and progress reporting functionality

# Import constants
. "$PSScriptRoot\Constants.ps1"

# Global variables for logging
$Global:LogFile = $null
$Global:ProgressCallback = $null

# Initialize logging
function Initialize-Logging {
    param(
        [string]$LogPath = "",
        [scriptblock]$ProgressCallback = $null
    )
    
    if ($LogPath -eq "") {
        $LogPath = Join-Path $env:TEMP "OfficeAddin-Setup-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
    }
    
    $Global:LogFile = $LogPath
    $Global:ProgressCallback = $ProgressCallback
    
    Write-LogMessage "INFO" "Logging initialized. Log file: $LogPath"
    return $LogPath
}

# Write log message with different levels
function Write-LogMessage {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG", "SUCCESS")]
        [string]$Level,
        
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [string]$Component = "MAIN"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] [$Component] $Message"
    
    # Write to console with colors
    switch ($Level) {
        "INFO" { Write-Host $logEntry -ForegroundColor White }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR" { Write-Host $logEntry -ForegroundColor Red }
        "DEBUG" { Write-Host $logEntry -ForegroundColor Gray }
    }
    
    # Write to log file if available
    if ($Global:LogFile) {
        try {
            Add-Content -Path $Global:LogFile -Value $logEntry -ErrorAction SilentlyContinue
        } catch {
            Write-Host "Failed to write to log file: $_" -ForegroundColor Red
        }
    }
}

# Report progress with status updates
function Report-Progress {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Status,
        
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [int]$PercentComplete = -1,
        
        [hashtable]$AdditionalData = @{}
    )
    
    $progressData = @{
        Status = $Status
        Message = $Message
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PercentComplete = $PercentComplete
        AdditionalData = $AdditionalData
    }
    
    Write-LogMessage "INFO" "Progress: $Status - $Message" "PROGRESS"
    
    # Call progress callback if available (for API updates)
    if ($Global:ProgressCallback) {
        try {
            & $Global:ProgressCallback $progressData
        } catch {
            Write-LogMessage "ERROR" "Failed to execute progress callback: $_" "PROGRESS"
        }
    }
    
    return $progressData
}

# Report error with structured data
function Report-Error {
    param(
        [Parameter(Mandatory=$true)]
        [int]$ErrorCode,
        
        [Parameter(Mandatory=$true)]
        [string]$ErrorMessage,
        
        [string]$Component = "UNKNOWN",
        
        [hashtable]$ErrorDetails = @{},
        
        [System.Management.Automation.ErrorRecord]$Exception = $null
    )
    
    $errorData = @{
        ErrorCode = $ErrorCode
        ErrorMessage = $ErrorMessage
        Component = $Component
        ErrorDetails = $ErrorDetails
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Exception = if ($Exception) { $Exception.ToString() } else { $null }
    }
    
    Write-LogMessage "ERROR" "[$ErrorCode] $ErrorMessage" $Component
    
    if ($Exception) {
        Write-LogMessage "DEBUG" "Exception details: $($Exception.ToString())" $Component
    }
    
    # Report as failed progress
    Report-Progress -Status (Get-StatusCode "FAILED") -Message $ErrorMessage -AdditionalData $errorData
    
    return $errorData
}

# Report success with structured data
function Report-Success {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [string]$Component = "MAIN",
        
        [hashtable]$SuccessData = @{}
    )
    
    Write-LogMessage "SUCCESS" $Message $Component
    
    $successData = @{
        Message = $Message
        Component = $Component
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        SuccessData = $SuccessData
    }
    
    return $successData
}

# Create structured result object
function New-Result {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$Success,
        
        [string]$Message = "",
        
        [int]$ErrorCode = 0,
        
        [hashtable]$Data = @{}
    )
    
    return @{
        Success = $Success
        Message = $Message
        ErrorCode = $ErrorCode
        Data = $Data
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
}

# Note: Export-ModuleMember is not needed for .ps1 files that are dot-sourced 