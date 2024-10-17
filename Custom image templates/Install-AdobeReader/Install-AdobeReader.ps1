<# 
    Script Name: Install-AdobeReader.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@it-ed.com
    Date:        October 17, 2023
    Description: This script automates the download and installation of Adobe Reader on Windows 10.
                 It validates the installation and logs the process to C:\ited\logs\adobe.log.
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
$logFile = "$logDirectory\adobe.log"

# Installer download URL (replace with the actual URL from Adobe)
$installerUrl = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2400320112/AcroRdrDC2400320112_fr_FR.exe"

# Path to save the downloaded installer
$installerPath = "C:\Temp\AdobeReaderInstaller.exe"

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
    Write-Log "Starting Adobe Reader installation script."

    # Create Temp directory if it doesn't exist
    if (!(Test-Path -Path "C:\Temp")) {
        Write-Log "Creating C:\Temp directory."
        New-Item -ItemType Directory -Path "C:\Temp" -Force
    }

    # Download the Adobe Reader installer
    Write-Log "Downloading Adobe Reader installer from $installerUrl."
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -ErrorAction Stop

    # Verify if the installer was downloaded
    if (Test-Path -Path $installerPath) {
        Write-Log "Installer downloaded successfully to $installerPath."
    } else {
        Write-Log "Installer download failed." "ERROR"
        Throw "Installer download failed."
    }

    # Install Adobe Reader silently
    Write-Log "Starting silent installation of Adobe Reader."
    $installArgs = "/sAll /msi /norestart /quiet"
    $process = Start-Process -FilePath $installerPath -ArgumentList $installArgs -Wait -PassThru

    # Check the exit code of the installer
    if ($process.ExitCode -eq 0) {
        Write-Log "Adobe Reader installed successfully."
    } else {
        Write-Log "Adobe Reader installation failed with exit code $($process.ExitCode)." "ERROR"
        Throw "Adobe Reader installation failed with exit code $($process.ExitCode)."
    }

    # Validate the installation by checking if Adobe Reader is installed
    Write-Log "Validating Adobe Reader installation."
    $app = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "Adobe Acrobat Reader*" }
    if ($app) {
        Write-Log "Adobe Reader is installed. Version: $($app.Version)"
    } else {
        Write-Log "Adobe Reader is not installed." "ERROR"
        Throw "Adobe Reader is not installed."
    }

    # Cleanup: Remove the installer file
    Write-Log "Cleaning up installer file."
    Remove-Item -Path $installerPath -Force -ErrorAction SilentlyContinue

    Write-Log "Adobe Reader installation script completed successfully."
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
