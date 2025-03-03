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
$registryPath = "HKCU:\Software\Classes\wordaddin"

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

Write-Host "Custom URI protocol 'wordaddin://' has been registered successfully!"
Write-Host "It will execute: $commandValue" 