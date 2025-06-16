param (
    [Parameter(Mandatory=$true)]
    [string]$ExePath
)

# Ensure the executable path exists
if (-not (Test-Path $ExePath)) {
    Write-Error "Executable not found at: $ExePath"
    exit 1
}

# Register the custom URI scheme
$registryPath = "HKLM:\Software\Classes\wordaddin"

# Verify we have admin rights before proceeding
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "Administrator privileges are required to register the URI handler at the machine level."
    exit 1
}

# Create the main key
if (-not (Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
}

# Set default value
Set-ItemProperty -Path $registryPath -Name "(Default)" -Value "URL:Word Add-in Protocol"

# Add URL Protocol value
Set-ItemProperty -Path $registryPath -Name "URL Protocol" -Value ""

# Create command structure
$shellPath = "$registryPath\shell\open\command"
if (-not (Test-Path $shellPath)) {
    New-Item -Path $shellPath -Force | Out-Null
}

# Set command to run our executable with the URI as parameter
$commandValue = "`"$ExePath`" `"%1`""
Set-ItemProperty -Path $shellPath -Name "(Default)" -Value $commandValue

Write-Host "Custom URI protocol 'wordaddin://' has been registered successfully for all users!"
Write-Host "It will execute: $commandValue"