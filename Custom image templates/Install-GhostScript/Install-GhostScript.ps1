<# 
    Script Name: Install-GhostScript.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@it-ed.com
    Date:        October 18, 2023
    Description: This script automates the download and installation of Ghostscript on Windows 10.
                 It validates the installation and logs the process to C:\ited\logs\ghostscript.log.
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
$logFile = "$logDirectory\ghostscript.log"

# URL for the Ghostscript installer
$ghostscriptUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-GhostScript/gs10021w64.msi"

# Path to save the downloaded installer
$ghostscriptInstallerPath = "C:\Temp\gs10021w64.msi"

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
    Write-Log "Starting Ghostscript installation script."

    # Create Temp directory if it doesn't exist
    if (!(Test-Path -Path "C:\Temp")) {
        Write-Log "Creating C:\Temp directory."
        New-Item -ItemType Directory -Path "C:\Temp" -Force
    }

    # Download the Ghostscript installer
    Write-Log "Downloading Ghostscript installer from $ghostscriptUrl."
    Invoke-WebRequest -Uri $ghostscriptUrl -OutFile $ghostscriptInstallerPath -ErrorAction Stop

    # Verify if the installer was downloaded
    if (Test-Path -Path $ghostscriptInstallerPath) {
        Write-Log "Ghostscript installer downloaded successfully to $ghostscriptInstallerPath."
    } else {
        Write-Log "Ghostscript installer download failed." "ERROR"
        Throw "Ghostscript installer download failed."
    }

    # Install Ghostscript silently
    Write-Log "Starting silent installation of Ghostscript."
    $ghostscriptInstallArgs = "/qn"
    $ghostscriptProcess = Start-Process -FilePath $ghostscriptInstallerPath -ArgumentList $ghostscriptInstallArgs -Wait -PassThru

    # Check the exit code of the Ghostscript installer
    if ($ghostscriptProcess.ExitCode -eq 0) {
        Write-Log "Ghostscript installed successfully."
    } else {
        Write-Log "Ghostscript installation failed with exit code $($ghostscriptProcess.ExitCode)." "ERROR"
        Throw "Ghostscript installation failed with exit code $($ghostscriptProcess.ExitCode)."
    }

    # Cleanup: Remove the installer file
    Write-Log "Cleaning up installer file."
    Remove-Item -Path $ghostscriptInstallerPath -Force -ErrorAction SilentlyContinue

    Write-Log "Ghostscript installation script completed successfully."
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
