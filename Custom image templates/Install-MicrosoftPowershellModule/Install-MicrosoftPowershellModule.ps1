<#
===============================================================================
                            Azure Custom Image Builder Script
===============================================================================
Creator      : Marc-Andre Brochu
Email        : mabrochu@it-ed.com
Date         : 12 Novembre 2024
Description  : This script installs all PowerShell modules for Office 365 and Azure
               on a custom Azure image. The installation process logs all activities
               into C:\ited\logs with a specific prefix for easy identification.
               This installation will be available for all users on the system.
===============================================================================
#>

# Define paths for logs and temporary files
$logDir = "C:\ited\logs"
$Date = Get-Date -Format "yyyyMMdd-HHmmss"
$logFile = "$logDir\AIB-ModuleInstall-$Date.log"

# Function to write logs
function Write-Log {
    param($Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "$TimeStamp - $Message"
    Add-Content -Path $logFile -Value $LogMessage
    Write-Output $LogMessage
}

# Create directories if they do not exist
if (-not (Test-Path -Path $logDir)) {
    try {
        New-Item -ItemType Directory -Path $logDir -Force
        Write-Log "Log directory created: $logDir"
    }
    catch {
        Write-Error "Unable to create log directory: $_"
        exit 1
    }
}

Write-Log "Starting installation of Office 365 and Azure PowerShell modules..."

# Install NuGet provider silently
Write-Log "Installing NuGet provider silently..."
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

# Check if NuGet is already installed
$nugetProvider = Get-PackageProvider -ListAvailable | Where-Object { $_.Name -eq 'NuGet' }
if (-not $nugetProvider) {
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers
        Write-Log "NuGet installed successfully"
    }
    catch {
        Write-Log "ERROR: Unable to install NuGet: $_"
        exit 1
    }
} else {
    Write-Log "NuGet is already installed (Version: $($nugetProvider.Version))"
}

# Configure PSGallery as a trusted repository
try {
    Write-Log "Configuring PSGallery as a trusted repository..."
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Write-Log "PSGallery configured successfully"
}
catch {
    Write-Log "ERROR: Unable to configure PSGallery: $_"
    exit 1
}

# List of modules to install
$ModulesToInstall = @(
    'Az',
    'MSOnline',
    'AzureAD',
    'Microsoft.Online.SharePoint.PowerShell',
    'ExchangeOnlineManagement',
    'Microsoft.Graph',
    'Microsoft.PowerApps.PowerShell',
    'Microsoft.PowerApps.Administration.PowerShell',
    'SharePointPnPPowerShellOnline',
    'MicrosoftTeams'
)

# Install modules for all users
foreach ($Module in $ModulesToInstall) {
    try {
        Write-Log "Installing module: $Module"
        $InstallParams = @{
            Name               = $Module
            Force              = $true
            AllowClobber       = $true
            Scope              = 'AllUsers'
            Repository         = 'PSGallery'
            SkipPublisherCheck = $true
            Confirm            = $false
        }
        
        Install-Module @InstallParams
        Write-Log "Module $Module installed successfully"
        
        # Verify installation
        $InstalledModule = Get-Module -Name $Module -ListAvailable
        Write-Log "Installed version of $Module $($InstalledModule.Version)"
    }
    catch {
        Write-Log "ERROR: Failed to install module $Module $_"
        # Continue with the next module instead of stopping the script
        continue
    }
}

# Final verification
Write-Log "Verifying installed modules..."
foreach ($Module in $ModulesToInstall) {
    $InstalledModule = Get-Module -Name $Module -ListAvailable
    if ($InstalledModule) {
        Write-Log "✓ $Module is installed (Version: $($InstalledModule.Version))"
    } else {
        Write-Log "⚠ $Module was not installed correctly"
    }
}

Write-Log "Installation complete"
