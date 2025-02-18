# Define the registry path
$registryPath = "HKCU:\Software\Microsoft\Office\Word\Addins\my-word-add-in"

# Create the registry key for the add-in
New-Item -Path $registryPath -Force

# Set the required registry values
New-ItemProperty -Path $registryPath -Name "Description" -Value "My Word Add-in Description" -PropertyType String -Force
New-ItemProperty -Path $registryPath -Name "FriendlyName" -Value "My Word Add-in" -PropertyType String -Force
New-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force
New-ItemProperty -Path $registryPath -Name "Manifest" -Value "file:///D:/Task/Working_On/my-word-add-in-app/word-add-in/manifest.xml" -PropertyType String -Force

Write-Output "Add-in installed and enabled successfully."
