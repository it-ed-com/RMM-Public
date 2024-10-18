<#
    Script Name: Install-Office.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@it-ed.com
    Date:        October 17, 2023
    Description: This script automates the download and installation of Office 365 Shared Computer Licensing.
                 It modifies the Configuration.xml, installs Office using setup.exe, and logs all actions.
#>

# =========================================
# Configuration Parameters
# =========================================

# Log directory
$logDirectory = "C:\ited\logs"
if (!(Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

# Log file path
$logFile = "$logDirectory\office_install.log"

# Download URLs
$setupUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-Office365-French/setup.exe"
$configUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-Office365-French/Configuration.xml"

# Download paths
$officeDir = "C:\ited\Office365"
$setupPath = "$officeDir\setup.exe"
$configPath = "$officeDir\Configuration.xml"
$modifiedConfigPath = "$officeDir\Configuration_modified.xml"

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
    Add-Content -Path $logFile -Value $logMessage
    Write-Host $logMessage
}

# =========================================
# XML Modification Function
# =========================================

Function Modify-ConfigurationXml {
    param (
        [string]$sourcePath,
        [string]$destinationPath
    )
    
    Write-Log "Modifying Configuration XML."

    try {
        # Load the XML
        [xml]$xml = Get-Content -Path $sourcePath

        # Replace all occurrences of "HelpOX" with "ited" in SourcePath
        $originalSourcePath = $xml.Configuration.Add.SourcePath
        $modifiedSourcePath = $originalSourcePath -replace "HelpOX", "ited"
        $xml.Configuration.Add.SourcePath = $modifiedSourcePath
        Write-Log "SourcePath changed from '$originalSourcePath' to '$modifiedSourcePath'."

        # Change language from "fr-fr" to "fr-ca"
        foreach ($lang in $xml.Configuration.Add.Product.Language) {
            if ($lang.ID -eq "fr-fr") {
                $lang.ID = "fr-ca"
                Write-Log "Language changed from 'fr-fr' to 'fr-ca'."
            }
        }

        # Change language in LanguagePack
        $xml.Configuration.Add.Product.LanguagePack.Language.ID = "fr-ca"
        Write-Log "LanguagePack changed from 'fr-fr' to 'fr-ca'."

        # Exclude all applications except Word and Excel
        $appsToExclude = @("Access", "Groove", "Lync", "OneDrive", "OneNote", "Outlook", "PowerPoint", "Publisher", "Teams")
        foreach ($app in $appsToExclude) {
            # Check if ExcludeApp already exists to avoid duplication
            $exists = $xml.Configuration.Add.Product.ExcludeApp | Where-Object { $_.ID -eq $app }
            if (-not $exists) {
                $excludeAppNode = $xml.CreateElement("ExcludeApp")
                $excludeAppNode.SetAttribute("ID", $app)
                $xml.Configuration.Add.Product.AppendChild($excludeAppNode) | Out-Null
                Write-Log "Excluded application: $app."
            }
        }

        # Save the modified XML
        $xml.Save($destinationPath)
        Write-Log "Modified Configuration XML saved to '$destinationPath'."
    }
    catch {
        Write-Log "Error modifying Configuration XML: $_" "ERROR"
        Throw "Configuration XML modification failed."
    }
}

# =========================================
# Installation Validation Function
# =========================================

Function Is-OfficeInstalled {
    param (
        [string]$productName
    )
    $officeKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot"
    )
    
    foreach ($key in $officeKeys) {
        if (Test-Path -Path $key) {
            $installed = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
            if ($installed -and $installed.InstallPath) {
                $officePath = $installed.InstallPath
                $productExe = switch ($productName) {
                    "Word" { "WINWORD.EXE" }
                    "Excel" { "EXCEL.EXE" }
                    default { "" }
                }
                if ($productExe -ne "") {
                    $exePath = Join-Path -Path $officePath -ChildPath $productExe
                    if (Test-Path -Path $exePath) {
                        return $true
                    }
                }
            }
        }
    }
    return $false
}

# =========================================
# Script Execution
# =========================================

Try {
    Write-Log "Starting Office 365 Shared Computer Licensing installation script."

    # Ensure the script is running as Administrator
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")) {
        Write-Log "Script is not running with administrative privileges." "ERROR"
        Throw "Please run this script as an Administrator."
    }

    # Create Office365 directory if it doesn't exist
    if (!(Test-Path -Path $officeDir)) {
        Write-Log "Creating directory '$officeDir'."
        New-Item -ItemType Directory -Path $officeDir -Force | Out-Null
    }

    # Download setup.exe
    Write-Log "Downloading setup.exe from '$setupUrl'."
    Invoke-WebRequest -Uri $setupUrl -OutFile $setupPath -ErrorAction Stop
    Write-Log "setup.exe downloaded successfully to '$setupPath'."

    # Download Configuration.xml
    Write-Log "Downloading Configuration.xml from '$configUrl'."
    Invoke-WebRequest -Uri $configUrl -OutFile $configPath -ErrorAction Stop
    Write-Log "Configuration.xml downloaded successfully to '$configPath'."

    # Modify Configuration.xml
    Modify-ConfigurationXml -sourcePath $configPath -destinationPath $modifiedConfigPath

    # Download Office sources
    Write-Log "Starting Office download using modified Configuration.xml."
    Start-Process -FilePath $setupPath -ArgumentList "/download `"$modifiedConfigPath`"" -Wait -NoNewWindow -ErrorAction Stop
    Write-Log "Office download completed."

    # Install Office using modified Configuration.xml
    Write-Log "Starting Office installation using modified Configuration.xml."
    Start-Process -FilePath $setupPath -ArgumentList "/configure `"$modifiedConfigPath`"" -Wait -NoNewWindow -ErrorAction Stop
    Write-Log "Office installation initiated."

    # Validate installation of Word and Excel
    Write-Log "Validating Office installation for Word and Excel."
    $isWordInstalled = Is-OfficeInstalled -productName "Word"
    $isExcelInstalled = Is-OfficeInstalled -productName "Excel"

    if ($isWordInstalled -and $isExcelInstalled) {
        Write-Log "Office installed successfully: Word and Excel are available."
    }
    else {
        Write-Log "Office installation failed: Word and/or Excel are not installed." "ERROR"
        Throw "Office installation validation failed."
    }

    # Cleanup downloaded files
    Write-Log "Cleaning up downloaded files."
    Remove-Item -Path $setupPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $modifiedConfigPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $configPath -Force -ErrorAction SilentlyContinue
    Write-Log "Downloaded files have been removed."

    Write-Log "Office 365 Shared Computer Licensing installation script completed successfully."
    Exit 0
}
Catch {
    Write-Log "An error occurred: $_" "ERROR"
    Exit 1
}
Finally {
    # Nothing to do here
}
