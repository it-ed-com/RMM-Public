<#
===============================================================================
                            ConnectWise Manage Silent Installer
===============================================================================
Creator      : Marc-Andre Brochu
Email        : mabrochu@it-ed.com
Date         : 13 Novembre 2024
Description  : This script installs ConnectWise Manage silently using an MSI
               downloaded from an external URL. The installation process logs
               all activities into C:\ited\logs with a specific prefix for easy
               identification. This installation will be available for all users
               on the system.
===============================================================================
#>

# Define paths for logs and temporary files
$logDir = "C:\ited\logs"
$tempDir = "C:\ited\temp"
$Date = Get-Date -Format "yyyyMMdd-HHmmss"
$logFile = "$logDir\CW-Manage-Install-$Date.log"
$msiPath = "$tempDir\ConnectWise-PSA-Internet-Client-x64.msi"

# URL for the MSI download
$msiUrl = "https://university.connectwise.com/install/2024/ConnectWise-PSA-Internet-Client-x64.msi"

# Function to write logs
function Write-Log {
    param($Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "$TimeStamp - $Message"
    try {
        Add-Content -Path $logFile -Value $LogMessage
    } catch {
        Write-Output "Unable to write to log file: $_"
    }
    Write-Output $LogMessage
}

# Create directories if they do not exist
if (-not (Test-Path -Path $tempDir)) {
    try {
        New-Item -ItemType Directory -Path $tempDir -Force
        Write-Log "Temp directory created: $tempDir"
    }
    catch {
        Write-Error "Unable to create temp directory: $_"
        exit 1
    }
}

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

Write-Log "Starting ConnectWise Manage installation..."

# Download the MSI from external URL
try {
    Write-Log "Downloading ConnectWise Manage MSI..."
    Invoke-WebRequest -Uri $msiUrl -OutFile $msiPath -UseBasicParsing
    Write-Log "Download completed: $msiPath"
}
catch {
    Write-Log "ERROR: Failed to download the MSI: $_"
    exit 1
}

# Install the MSI silently
try {
    Write-Log "Installing ConnectWise Manage silently..."
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$msiPath`" /q /norestart" -Wait
    Write-Log "ConnectWise Manage installed successfully"
}
catch {
    Write-Log "ERROR: Installation failed: $_"
    exit 1
}

# Finalize
Write-Log "ConnectWise Manage installation script completed."
