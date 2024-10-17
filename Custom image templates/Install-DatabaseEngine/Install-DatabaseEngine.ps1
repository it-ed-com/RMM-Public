<#
    Script Name: Install-AccessDatabaseEngine.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@it-ed.com
    Date:        October 17, 2023
    Description: This script automates the download and installation of Access Database Engine on Windows 10.
                 It validates the installation and logs the process to C:\ited\logs\databaseengine.log.
#>

# =========================================
# Configuration Parameters
# =========================================

# Ensure the log directory exists
$logDirectory = "C:\ited\logs"
if (!(Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force
}

# Log file path
$logFile = "$logDirectory\databaseengine.log"

# Installer download URL (use the URL provided)
$installerUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-DatabaseEngine/AccessDatabaseEngine.exe"

# Path to save the downloaded installer
$installerPath = "C:\Temp\AccessDatabaseEngine.exe"

# =========================================
# Logging Functions
# =========================================

Function Write-Log {
    param (
        [string]$message,
        [string]$type = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$type] $message"
    Write-Host $logMessage
}

# Start transcript to log all output to the log file
Start-Transcript -Path $logFile -Append -NoClobber

# =========================================
# Script Execution
# =========================================

Try {
    Write-Log "Starting Access Database Engine installation script."

    # Create Temp directory if it doesn't exist
    if (!(Test-Path -Path "C:\Temp")) {
        Write-Log "Creating C:\Temp directory."
        New-Item -ItemType Directory -Path "C:\Temp" -Force
    }

    # Download the Access Database Engine installer
    Write-Log "Downloading Access Database Engine installer from $installerUrl."
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -ErrorAction Stop

    # Verify if the installer was downloaded
    if (Test-Path -Path $installerPath) {
        Write-Log "Installer downloaded successfully to $installerPath."
    } else {
        Write-Log "Installer download failed." "ERROR"
        Throw "Installer download failed."
    }

    # Install Access Database Engine silently without restart
    Write-Log "Starting silent installation of Access Database Engine without restart."
    $installArgs = "/quiet /norestart"
    $process = Start-Process -FilePath $installerPath -ArgumentList $installArgs -Wait -PassThru

    # Check the exit code of the installer
    if ($process.ExitCode -eq 0) {
        Write-Log "Access Database Engine installed successfully."
    } else {
        Write-Log "Access Database Engine installation failed with exit code $($process.ExitCode)." "ERROR"
        Throw "Access Database Engine installation failed with exit code $($process.ExitCode)."
    }

    # Validate the installation
    Write-Log "Validating Access Database Engine installation."
    $installedProduct = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "Microsoft Access database engine*" }
    if ($installedProduct) {
        Write-Log "Access Database Engine is installed. Version: $($installedProduct.Version)"
    } else {
        Write-Log "Access Database Engine is not installed." "ERROR"
        Throw "Access Database Engine is not installed."
    }

    # Cleanup: Remove the installer file
    Write-Log "Cleaning up installer file."
    Remove-Item -Path $installerPath -Force -ErrorAction SilentlyContinue

    Write-Log "Access Database Engine installation script completed successfully."
    Exit 0
}
Catch {
    Write-Log "An error occurred: $_" "ERROR"
    Exit 1
}
Finally {
    # Stop transcript to finalize the log file
    Stop-Transcript
}
