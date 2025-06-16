# Office Add-in Setup - Constants and Configuration
# This file contains all global constants, error codes, and configuration values

# Error code constants
$Global:ERROR_CODES = @{
    WORD_NOT_INSTALLED = 100
    MANIFEST_CREATION_FAILED = 101
    NETWORK_SHARE_FAILED = 102
    OPEN_DOCUMENT_FAILED = 103
    ADDIN_CONFIGURATION_FAILED = 104
    USER_NOT_FOUND = 105
    INSUFFICIENT_PERMISSIONS = 106
    REGISTRY_ACCESS_FAILED = 107
    FILE_SYSTEM_ACCESS_FAILED = 108
    COM_OBJECT_FAILED = 109
    UI_AUTOMATION_FAILED = 110
}

# Error descriptions
$Global:ERROR_DESCRIPTIONS = @{
    100 = "Microsoft Word is not installed on this device."
    101 = "Failed to create or verify the manifest file."
    102 = "Failed to create or verify network share."
    103 = "Failed to open the specified document."
    104 = "Failed to configure the Office add-in."
    105 = "Could not determine current user context."
    106 = "Insufficient permissions to perform operation."
    107 = "Failed to access or modify registry."
    108 = "Failed to access file system."
    109 = "Failed to create or use COM object."
    110 = "Failed to perform UI automation."
}

# Status codes for progress reporting
$Global:STATUS_CODES = @{
    STARTING = "STARTING"
    CHECKING_WORD = "CHECKING_WORD"
    CREATING_MANIFEST = "CREATING_MANIFEST"
    CONFIGURING_SHARE = "CONFIGURING_SHARE"
    SETTING_TRUST = "SETTING_TRUST"
    OPENING_DOCUMENT = "OPENING_DOCUMENT"
    CONFIGURING_ADDIN = "CONFIGURING_ADDIN"
    COMPLETED = "COMPLETED"
    FAILED = "FAILED"
}

# Configuration constants
$Global:CONFIG = @{
    OFFICE_VERSION = "16.0"
    REGISTRY_BASE_PATH = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
    MANIFEST_FILENAME = "manifest.xml"
    DEFAULT_SHARE_DESCRIPTION = "Office Add-ins Shared Folder"
    TIMEOUT_SECONDS = 30
    UI_WAIT_MILLISECONDS = 500
    PROCESS_WAIT_SECONDS = 3
}

# Add-in configuration
$Global:ADDIN_CONFIG = @{
    ID = "f85491a7-0cf8-4950-b18c-d85ae9970d61"
    NAME = "IP Agent AI"
    VERSION = "1.0.0.0"
    PROVIDER = "AnyGenAI"
    BASE_URL = "http://10.100.100.71:3002"
    PERMISSIONS = "ReadWriteDocument"
}

# Export functions to make constants available globally
function Get-ErrorCode { param($name) return $Global:ERROR_CODES[$name] }
function Get-ErrorDescription { param($code) return $Global:ERROR_DESCRIPTIONS[$code] }
function Get-StatusCode { param($name) return $Global:STATUS_CODES[$name] }
function Get-Config { param($name) return $Global:CONFIG[$name] }
function Get-AddinConfig { param($name) return $Global:ADDIN_CONFIG[$name] }

# Note: Export-ModuleMember is not needed for .ps1 files that are dot-sourced 