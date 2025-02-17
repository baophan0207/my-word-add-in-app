Write-Host "Cleaning up and reinstalling Word Add-in..."

# Clear Office Cache
$wefPath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (Test-Path $wefPath) {
    Remove-Item -Path $wefPath -Recurse -Force
    Write-Host "✅ Cleared Office Cache" -ForegroundColor Green
}

# Remove existing registry entries
$regPaths = @(
    "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer",
    "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedProtocols\https\localhost"
)

foreach ($path in $regPaths) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Recurse -Force
        Write-Host "✅ Removed $path" -ForegroundColor Green
    }
}

Write-Host "`nPlease:"
Write-Host "1. Close all Word instances"
Write-Host "2. Run the add-in installation again from your application"
Write-Host "3. Start Word and check the ribbon"
Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 