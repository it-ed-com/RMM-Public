<#
    Script Name: Install-Fidelio.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@it-ed.com
    Date:        October 17, 2023
    Requirment:  .NET Framework version 2.0
    Description: This script automates the download and installation of Fidelio on Windows 10.
                 It validates the installation, removes the desktop icon, and logs the process to C:\ited\logs\fidelio.log.
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
$logFile = "$logDirectory\fidelio.log"

# Installer download URL
$installerUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-Fidelio/FidelioSetup.msi"

# Path to save the downloaded installer
$installerPath = "C:\Temp\FidelioSetup.msi"

# Installation directory
$installDir = "C:\Fidelio"

# =========================================
# Logging Function
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
    Write-Log "Starting Fidelio installation script."

    # Create Temp directory if it doesn't exist
    if (!(Test-Path -Path "C:\Temp")) {
        Write-Log "Creating C:\Temp directory."
        New-Item -ItemType Directory -Path "C:\Temp" -Force
    }

    # Download the Fidelio installer
    Write-Log "Downloading Fidelio installer from $installerUrl."
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -ErrorAction Stop

    # Verify if the installer was downloaded
    if (Test-Path -Path $installerPath) {
        Write-Log "Installer downloaded successfully to $installerPath."
    } else {
        Write-Log "Installer download failed." "ERROR"
        Throw "Installer download failed."
    }

    # Install Fidelio silently with specified arguments
    Write-Log "Starting silent installation of Fidelio."
    $installArgs = "/i `"$installerPath`" TARGETDIR=`"$installDir`" /q"
    $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $installArgs -Wait -PassThru

    # Check the exit code of the installer
    if ($process.ExitCode -eq 0) {
        Write-Log "Fidelio installed successfully."
    } else {
        Write-Log "Fidelio installation failed with exit code $($process.ExitCode)." "ERROR"
        Throw "Fidelio installation failed with exit code $($process.ExitCode)."
    }

    # Validate the installation
    Write-Log "Validating Fidelio installation."
    if (Test-Path -Path "$installDir\Fidelio.ico") {
        Write-Log "Fidelio is installed at $installDir."
    } else {
        Write-Log "Fidelio is not installed in $installDir." "ERROR"
        Throw "Fidelio is not installed in $installDir."
    }

    # Remove desktop icons
    Write-Log "Removing desktop icons."
    Remove-Item "C:\Users\Public\Desktop\*" -Force -ErrorAction SilentlyContinue

    # Cleanup: Remove the installer file
    Write-Log "Cleaning up installer file."
    Remove-Item -Path $installerPath -Force -ErrorAction SilentlyContinue

    Write-Log "Fidelio installation script completed successfully."
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
