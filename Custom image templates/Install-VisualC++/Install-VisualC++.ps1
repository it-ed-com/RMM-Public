<#
    Script Name: Install-VisualC++.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@it-ed.com
    Date:        October 17, 2023
    Description: This script automates the download and installation of Visual C++ Redistributable 2015-2022 (x64 and x86).
                 It validates each installation before proceeding to the next and logs the process to C:\ited\logs\visualc.log.
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
$logFile = "$logDirectory\visualc.log"

# Installer download URLs
$x64InstallerUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-VisualC%2B%2B/VC_redist.x64.exe"
$x86InstallerUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-VisualC%2B%2B/VC_redist.x86.exe"

# Paths to save the downloaded installers
$installerPathX64 = "C:\Temp\VC_redist.x64.exe"
$installerPathX86 = "C:\Temp\VC_redist.x86.exe"

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

# =========================================
# Validation Function
# =========================================

Function Is-VcRedistInstalled {
    param (
        [string]$architecture
    )
    
    # Déterminer si le système est 64 bits
    $is64Bit = [Environment]::Is64BitOperatingSystem

    if ($architecture -eq "x86" -and $is64Bit) {
        # Sur un OS 64 bits, les redistribuables x86 sont sous WOW6432Node
        $key = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x86"
    }
    elseif ($architecture -eq "x64") {
        # Pour les redistribuables x64
        $key = "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64"
    }
    elseif ($architecture -eq "x86" -and -not $is64Bit) {
        # Sur un OS 32 bits, les redistribuables x86 sont sous SOFTWARE\Microsoft
        $key = "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x86"
    }
    else {
        Write-Log "Architecture non prise en charge pour la validation : $architecture" "ERROR"
        return $false
    }

    if (Test-Path -Path $key) {
        $installed = (Get-ItemProperty -Path $key).Installed
        return $installed -eq 1
    }
    else {
        return $false
    }
}

# =========================================
# Script Execution
# =========================================

Try {
    Write-Log "Starting Visual C++ Redistributable installation script."

    # Ensure script is running as Administrator
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")) {
        Write-Log "Script is not running with administrative privileges." "ERROR"
        Throw "Please run this script as an Administrator."
    }

    # Create Temp directory if it doesn't exist
    if (!(Test-Path -Path "C:\Temp")) {
        Write-Log "Creating C:\Temp directory."
        New-Item -ItemType Directory -Path "C:\Temp" -Force
    }

    # Start transcript to log all output to the log file
    Start-Transcript -Path $logFile -Append -NoClobber

    # Download the Visual C++ Redistributable x64 installer
    Write-Log "Downloading Visual C++ Redistributable x64 from $x64InstallerUrl."
    Invoke-WebRequest -Uri $x64InstallerUrl -OutFile $installerPathX64 -ErrorAction Stop

    # Verify if the x64 installer was downloaded
    if (Test-Path -Path $installerPathX64) {
        Write-Log "Installer x64 downloaded successfully."
    }
    else {
        Write-Log "Installer x64 failed to download." "ERROR"
        Throw "Installer x64 download failed."
    }

    # Install Visual C++ Redistributable x64 silently
    Write-Log "Starting silent installation of Visual C++ Redistributable x64."
    $installArgsX64 = "/install /quiet /norestart"
    $processX64 = Start-Process -FilePath $installerPathX64 -ArgumentList $installArgsX64 -Wait -PassThru

    # Check the exit code of the x64 installer
    if ($processX64.ExitCode -eq 0) {
        Write-Log "Visual C++ Redistributable x64 installed successfully."
    }
    else {
        Write-Log "Visual C++ Redistributable x64 installation failed with exit code $($processX64.ExitCode)." "ERROR"
        Throw "Visual C++ Redistributable x64 installation failed with exit code $($processX64.ExitCode)."
    }

    # Validate the x64 installation by checking the registry
    Write-Log "Validating Visual C++ Redistributable x64 installation."

    $isX64Installed = Is-VcRedistInstalled -architecture "x64"

    if ($isX64Installed) {
        Write-Log "Visual C++ Redistributable x64 is installed."
    }
    else {
        Write-Log "Visual C++ Redistributable x64 is not installed." "ERROR"
        Throw "Validation failed for Visual C++ Redistributable x64."
    }

    # Cleanup: Remove the x64 installer file
    Write-Log "Cleaning up x64 installer file."
    Remove-Item -Path $installerPathX64 -Force -ErrorAction SilentlyContinue

    # Download the Visual C++ Redistributable x86 installer
    Write-Log "Downloading Visual C++ Redistributable x86 from $x86InstallerUrl."
    Invoke-WebRequest -Uri $x86InstallerUrl -OutFile $installerPathX86 -ErrorAction Stop

    # Verify if the x86 installer was downloaded
    if (Test-Path -Path $installerPathX86) {
        Write-Log "Installer x86 downloaded successfully."
    }
    else {
        Write-Log "Installer x86 failed to download." "ERROR"
        Throw "Installer x86 download failed."
    }

    # Install Visual C++ Redistributable x86 silently
    Write-Log "Starting silent installation of Visual C++ Redistributable x86."
    $installArgsX86 = "/install /quiet /norestart"
    $processX86 = Start-Process -FilePath $installerPathX86 -ArgumentList $installArgsX86 -Wait -PassThru

    # Check the exit code of the x86 installer
    if ($processX86.ExitCode -eq 0) {
        Write-Log "Visual C++ Redistributable x86 installed successfully."
    }
    else {
        Write-Log "Visual C++ Redistributable x86 installation failed with exit code $($processX86.ExitCode)." "ERROR"
        Throw "Visual C++ Redistributable x86 installation failed with exit code $($processX86.ExitCode)."
    }

    # Validate the x86 installation by checking the registry
    Write-Log "Validating Visual C++ Redistributable x86 installation."

    $isX86Installed = Is-VcRedistInstalled -architecture "x86"

    if ($isX86Installed) {
        Write-Log "Visual C++ Redistributable x86 is installed."
    }
    else {
        Write-Log "Visual C++ Redistributable x86 is not installed." "ERROR"
        Throw "Validation failed for Visual C++ Redistributable x86."
    }

    # Cleanup: Remove the x86 installer file
    Write-Log "Cleaning up x86 installer file."
    Remove-Item -Path $installerPathX86 -Force -ErrorAction SilentlyContinue

    Write-Log "Visual C++ Redistributable installation script completed successfully."
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
