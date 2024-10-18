<#
    Script Name: Install-NetFramework35.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@it-ed.com
    Date:        October 17, 2023
    Description: This script automates the installation of .NET Framework 3.5 on Windows 10.
                 It validates the installation and logs the process to C:\ited\logs\netframework.log.
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
$logFile = "$logDirectory\netframework.log"

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
    Write-Log "Starting .NET Framework 3.5 installation script."

    # Ensure script is running as Administrator
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")) {
        Write-Log "Script is not running with administrative privileges." "ERROR"
        Throw "Please run this script as an Administrator."
    }

    # Check if .NET Framework 3.5 is already installed
    Write-Log "Checking if .NET Framework 3.5 is already installed."
    $net35 = Get-WindowsOptionalFeature -Online -FeatureName NetFx3
    if ($net35.State -eq "Enabled") {
        Write-Log ".NET Framework 3.5 is already installed."
        Exit 0
    }

    # Install .NET Framework 3.5 using DISM
    Write-Log "Installing .NET Framework 3.5 using DISM."
    $installProcess = Start-Process -FilePath "dism.exe" -ArgumentList "/Online", "/Enable-Feature", "/FeatureName:NetFx3", "/All", "/Quiet", "/NoRestart" -Wait -PassThru

    # Check the exit code of DISM
    if ($installProcess.ExitCode -eq 0) {
        Write-Log ".NET Framework 3.5 installed successfully."
    } else {
        Write-Log ".NET Framework 3.5 installation failed with exit code $($installProcess.ExitCode)." "ERROR"
        Throw ".NET Framework 3.5 installation failed with exit code $($installProcess.ExitCode)."
    }

    # Validate the installation
    Write-Log "Validating .NET Framework 3.5 installation."
    $net35 = Get-WindowsOptionalFeature -Online -FeatureName NetFx3
    if ($net35.State -eq "Enabled") {
        Write-Log ".NET Framework 3.5 is installed successfully."
    } else {
        Write-Log ".NET Framework 3.5 is not installed." "ERROR"
        Throw ".NET Framework 3.5 is not installed."
    }

    Write-Log ".NET Framework 3.5 installation script completed successfully."
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
