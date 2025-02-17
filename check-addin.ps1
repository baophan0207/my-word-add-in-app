Write-Host "Checking Word Add-in Installation Status..."

# Check manifest registration
$manifestPath = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Name "UserDevManifests" -ErrorAction SilentlyContinue

if ($manifestPath) {
    Write-Host "`n✅ Add-in manifest is registered:" -ForegroundColor Green
    Write-Host $manifestPath.UserDevManifests
} else {
    Write-Host "`n❌ Add-in manifest is not registered" -ForegroundColor Red
}

# Check if manifest file exists
if ($manifestPath) {
    if (Test-Path $manifestPath.UserDevManifests) {
        Write-Host "`n✅ Manifest file exists" -ForegroundColor Green
        Write-Host "Manifest content:"
        Get-Content $manifestPath.UserDevManifests
    } else {
        Write-Host "`n❌ Manifest file not found at specified path" -ForegroundColor Red
    }
}

# Check auto-run settings
$settings = @(
    @{Path="HKCU:\Software\Microsoft\Office\16.0\Word\Options"; Name="EnableAutoOpenTaskpane"},
    @{Path="HKCU:\Software\Microsoft\Office\16.0\WEF\Developer\Runtime\Settings"; Name="EnableAutoOpenTaskpane"},
    @{Path="HKCU:\Software\Microsoft\Office\16.0\Word\Options"; Name="TaskPanesAllowed"},
    @{Path="HKCU:\Software\Microsoft\Office\16.0\Word\Options"; Name="ShowTaskPane"}
)

Write-Host "`nChecking Registry Settings:" -ForegroundColor Yellow
foreach ($setting in $settings) {
    $value = Get-ItemProperty -Path $setting.Path -Name $setting.Name -ErrorAction SilentlyContinue
    if ($value) {
        Write-Host "✅ $($setting.Name): $($value.$($setting.Name))" -ForegroundColor Green
    } else {
        Write-Host "❌ $($setting.Name) not set" -ForegroundColor Red
    }
}

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")