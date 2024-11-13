<#
===============================================================================
                            ConnectWise Automate Silent Installer
===============================================================================
Creator      : Marc-Andre Brochu
Email        : mabrochu@it-ed.com
Date         : 13 Novembre 2024
Description  : This script installs ConnectWise Automate silently using an executable
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
$logFile = "$logDir\CW-Automate-Install-$Date.log"
$installerPath = "$tempDir\ControlCenterInstaller.exe"

# URL for the executable download
$installerUrl = "https://automate.cloudox.ca/LabTech/Updates/ControlCenterInstaller.exe"

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

Write-Log "Starting ConnectWise Automate installation..."

# Download the executable from external URL
try {
    Write-Log "Downloading ConnectWise Automate installer..."
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing -ErrorAction Stop
    Write-Log "Download completed successfully: $installerPath"
}
catch {
    Write-Log "ERROR: Failed to download the installer: $_"
    exit 1
}

# Verify if installer file exists after download
if (-not (Test-Path -Path $installerPath)) {
    Write-Log "ERROR: Installer file not found at path: $installerPath"
    exit 1
} else {
    Write-Log "Installer file verified at path: $installerPath"
}

# Install the executable silently
try {
    Write-Log "Installing ConnectWise Automate silently..."
    Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait -PassThru
    Write-Log "ConnectWise Automate installed successfully"
}
catch {
    Write-Log "ERROR: Installation failed: $_"
    exit 1
}

# Verify installation by checking existence in Programs
try {
    $installedApp = Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name LIKE 'ConnectWise Automate%'"
    if ($null -ne $installedApp) {
        Write-Log "ConnectWise Automate installation verified successfully."
    } else {
        Write-Log "ERROR: ConnectWise Automate was not found in the list of installed programs."
        exit 1
    }
}
catch {
    Write-Log "ERROR: Failed to verify the installation: $_"
    exit 1
}

# Finalize
Write-Log "ConnectWise Automate installation script completed."
