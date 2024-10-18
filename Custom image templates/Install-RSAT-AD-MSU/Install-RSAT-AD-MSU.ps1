<#
    Script Name: Install-RSAT-AD-DS-LDS.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@it-ed.com
    Date:        October 17, 2023
    Description: This script installs the RSAT feature required for dsa.msc (Active Directory Users and Computers) on Windows 10.
                 It validates the installation and logs the process to C:\ited\logs\rsat.log.
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
$logFile = "$logDirectory\rsat.log"

# Feature name for RSAT Active Directory Users and Computers
$featureName = "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"

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
    Write-Log "Starting RSAT Active Directory Users and Computers installation script."

    # Ensure script is running as Administrator
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")) {
        Write-Log "Script is not running with administrative privileges." "ERROR"
        Throw "Please run this script as an Administrator."
    }

    # Check Windows version
    $osVersion = (Get-CimInstance -ClassName Win32_OperatingSystem).Version
    Write-Log "Detected Windows version: $osVersion"

    # Install RSAT Active Directory Users and Computers
    Write-Log "Installing feature: $featureName"

    # Check if the feature is already installed
    $featureState = Get-WindowsCapability -Online -Name $featureName

    if ($featureState.State -eq "Installed") {
        Write-Log "Feature $featureName is already installed."
    } elseif ($featureState.State -eq "NotPresent") {
        Try {
            Add-WindowsCapability -Online -Name $featureName -ErrorAction Stop
            Write-Log "Feature $featureName installed successfully."
        }
        Catch {
            Write-Log "Failed to install feature $featureName. Error: $_" "ERROR"
            Throw "Failed to install feature $featureName."
        }
    } else {
        Write-Log "Feature $featureName is in state: $($featureState.State). Cannot proceed." "ERROR"
        Throw "Cannot install feature $featureName."
    }

    # Validate the installation
    Write-Log "Validating RSAT Active Directory Users and Computers installation."

    $featureState = Get-WindowsCapability -Online -Name $featureName

    if ($featureState.State -eq "Installed") {
        Write-Log "Feature $featureName is installed."
        Write-Log "RSAT Active Directory Users and Computers is installed successfully."
    } else {
        Write-Log "Feature $featureName is not installed." "ERROR"
        Throw "RSAT Active Directory Users and Computers installation failed."
    }

    Write-Log "RSAT Active Directory Users and Computers installation script completed successfully."
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
