<# 
    Script Name: Install-Office365.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@ited.com
    Date:        October 17, 2024
    Description: This script automates the download and installation of Office 365 in French Canada.
                 It downloads setup.exe and Configuration.xml from specified URLs,
                 modifies the Configuration.xml as needed, downloads Office sources,
                 installs Office using the configuration, validates the installation,
                 and logs the process to C:\ited\logs\Office365.log.
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
$logFile = "$logDirectory\Office365.log"

# URLs for the Office 365 installer and configuration
$setupUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-Office365-French/setup.exe"
$xmlUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-Office365-French/Configuration.xml"

# Paths to save the downloaded installer and configuration
$officeInstallerPath = "C:\ited\Office365\setup.exe"
$officeXMLPath = "C:\ited\Office365\Configuration.xml"

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
    Write-Log "Starting Office 365 installation script."

    # Create Office365 directory if it doesn't exist
    if (!(Test-Path -Path "C:\ited\Office365")) {
        Write-Log "Creating C:\ited\Office365 directory."
        New-Item -ItemType Directory -Path "C:\ited\Office365" -Force
    }

    # Download the Office 365 Setup installer
    Write-Log "Downloading Office 365 Setup installer from $setupUrl."
    Invoke-WebRequest -Uri $setupUrl -OutFile $officeInstallerPath -ErrorAction Stop

    # Verify if the installer was downloaded
    if (Test-Path -Path $officeInstallerPath) {
        Write-Log "Office 365 installer downloaded successfully to $officeInstallerPath."
    } else {
        Write-Log "Office 365 installer download failed." "ERROR"
        Throw "Office 365 installer download failed."
    }

    # Download the Office 365 configuration XML
    Write-Log "Downloading Office 365 configuration XML from $xmlUrl."
    Invoke-WebRequest -Uri $xmlUrl -OutFile $officeXMLPath -ErrorAction Stop

    # Verify if the configuration XML was downloaded
    if (Test-Path -Path $officeXMLPath) {
        Write-Log "Office 365 configuration XML downloaded successfully to $officeXMLPath."
    } else {
        Write-Log "Office 365 configuration XML download failed." "ERROR"
        Throw "Office 365 configuration XML download failed."
    }

    # Download Office 365 sources using the configuration file
    Write-Log "Downloading Office 365 sources using the configuration file."
    $downloadProcess = Start-Process -FilePath $officeInstallerPath -ArgumentList "/download `"$officeXMLPath`"" -Wait -PassThru

    # Check the exit code of the download process
    if ($downloadProcess.ExitCode -eq 0) {
        Write-Log "Office 365 sources downloaded successfully."
    } else {
        Write-Log "Office 365 sources download failed with exit code $($downloadProcess.ExitCode)." "ERROR"
        Throw "Office 365 sources download failed with exit code $($downloadProcess.ExitCode)."
    }

    # Install Office 365 using the configuration file
    Write-Log "Installing Office 365 using the configuration file."
    $installProcess = Start-Process -FilePath "C:\ited\Office365\setup.exe" -ArgumentList "/configure `"$officeXMLPath`"" -Wait -PassThru

    # Check the exit code of the installation process
    if ($installProcess.ExitCode -eq 0) {
        Write-Log "Office 365 installed successfully."
    } else {
        Write-Log "Office 365 installation failed with exit code $($installProcess.ExitCode)." "ERROR"
        Throw "Office 365 installation failed with exit code $($installProcess.ExitCode)."
    }

    # Validate the installation by checking if Office applications are installed
    Write-Log "Validating Office 365 installation."
    $officeProducts = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "Microsoft Office*" }
    if ($officeProducts) {
        Write-Log "Office 365 is installed successfully."
    } else {
        Write-Log "Office 365 installation validation failed." "ERROR"
        Throw "Office 365 installation validation failed."
    }

    # Cleanup: Remove the installer and configuration files
    Write-Log "Cleaning up installer and configuration files."
    Remove-Item -Path $officeInstallerPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $officeXMLPath -Force -ErrorAction SilentlyContinue

    Write-Log "Office 365 installation script completed successfully."
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
