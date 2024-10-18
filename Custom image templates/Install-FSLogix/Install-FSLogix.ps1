<# 
    Script Name: Install-FSLogix.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@it-ed.com
    Date:        October 18, 2023
    Description: This script automates the download and installation of FSLogix Apps on Windows.
                 It validates the installation and logs the process to C:\ited\logs\fslogix.log.
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
$logFile = "$logDirectory\fslogix.log"

# URL for the FSLogix installer
$fslogixUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-FSLogix/2.9.8884.27471/FSLogixAppsSetup.exe"

# Path to save the downloaded installer
$fslogixInstallerPath = "C:\Temp\FSLogixAppsSetup.exe"

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
    Write-Log "Starting FSLogix installation script."

    # Create Temp directory if it doesn't exist
    if (!(Test-Path -Path "C:\Temp")) {
        Write-Log "Creating C:\Temp directory."
        New-Item -ItemType Directory -Path "C:\Temp" -Force
    }

    # Download the FSLogix installer
    Write-Log "Downloading FSLogix installer from $fslogixUrl."
    Invoke-WebRequest -Uri $fslogixUrl -OutFile $fslogixInstallerPath -ErrorAction Stop

    # Verify if the installer was downloaded
    if (Test-Path -Path $fslogixInstallerPath) {
        Write-Log "FSLogix installer downloaded successfully to $fslogixInstallerPath."
    } else {
        Write-Log "FSLogix installer download failed." "ERROR"
        Throw "FSLogix installer download failed."
    }

    # Install FSLogix silently
    Write-Log "Starting silent installation of FSLogix."
    $fslogixInstallArgs = "/quiet"
    $fslogixProcess = Start-Process -FilePath $fslogixInstallerPath -ArgumentList $fslogixInstallArgs -Wait -PassThru

    # Check the exit code of the FSLogix installer
    if ($fslogixProcess.ExitCode -eq 0) {
        Write-Log "FSLogix installed successfully."
    } else {
        Write-Log "FSLogix installation failed with exit code $($fslogixProcess.ExitCode)." "ERROR"
        Throw "FSLogix installation failed with exit code $($fslogixProcess.ExitCode)."
    }

    # Cleanup: Remove the installer file
    Write-Log "Cleaning up installer file."
    Remove-Item -Path $fslogixInstallerPath -Force -ErrorAction SilentlyContinue

    Write-Log "FSLogix installation script completed successfully."
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
