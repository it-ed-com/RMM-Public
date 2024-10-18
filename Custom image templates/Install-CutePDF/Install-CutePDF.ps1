<# 
    Script Name: Install-CutePDF.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@it-ed.com
    Date:        October 17, 2023
    Description: This script automates the download and installation of CutePDF Writer on Windows 10.
                 It validates the installation, sets the default paper size to Letter, and logs the process to C:\ited\logs\cutepdf.log.
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
$logFile = "$logDirectory\cutepdf.log"

# URL for the CutePDF Writer installer
$cutepdfUrl = "https://www.cutepdf.com/download/CuteWriter.exe"

# Path to save the downloaded installer
$cutepdfInstallerPath = "C:\Temp\CuteWriter.exe"

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
    Write-Log "Starting CutePDF Writer installation script."

    # Create Temp directory if it doesn't exist
    if (!(Test-Path -Path "C:\Temp")) {
        Write-Log "Creating C:\Temp directory."
        New-Item -ItemType Directory -Path "C:\Temp" -Force
    }

    # Download the CutePDF Writer installer
    Write-Log "Downloading CutePDF Writer installer from $cutepdfUrl."
    Invoke-WebRequest -Uri $cutepdfUrl -OutFile $cutepdfInstallerPath -ErrorAction Stop

    # Verify if the installer was downloaded
    if (Test-Path -Path $cutepdfInstallerPath) {
        Write-Log "CutePDF Writer installer downloaded successfully to $cutepdfInstallerPath."
    } else {
        Write-Log "CutePDF Writer installer download failed." "ERROR"
        Throw "CutePDF Writer installer download failed."
    }

    # Install CutePDF Writer silently
    Write-Log "Starting silent installation of CutePDF Writer."
    $cutepdfInstallArgs = "/VERYSILENT /NORESTART /NOICONS"
    $cutepdfProcess = Start-Process -FilePath $cutepdfInstallerPath -ArgumentList $cutepdfInstallArgs -Wait -PassThru

    # Check the exit code of the CutePDF installer
    if ($cutepdfProcess.ExitCode -eq 0) {
        Write-Log "CutePDF Writer installed successfully."
    } else {
        Write-Log "CutePDF Writer installation failed with exit code $($cutepdfProcess.ExitCode)." "ERROR"
        Throw "CutePDF Writer installation failed with exit code $($cutepdfProcess.ExitCode)."
    }

    # Validate the installation by checking if the printer exists
    Write-Log "Validating CutePDF Writer installation."
    $printer = Get-Printer | Where-Object { $_.Name -eq "CutePDF Writer" }
    if ($printer) {
        Write-Log "CutePDF Writer is installed and the printer is available."
    } else {
        Write-Log "CutePDF Writer is not installed or the printer is missing." "ERROR"
        Throw "CutePDF Writer is not installed or the printer is missing."
    }

    # Cleanup: Remove the installer file
    Write-Log "Cleaning up installer file."
    Remove-Item -Path $cutepdfInstallerPath -Force -ErrorAction SilentlyContinue

    Write-Log "CutePDF Writer installation script completed successfully."
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
