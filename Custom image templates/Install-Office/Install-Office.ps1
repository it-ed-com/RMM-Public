<#
===============================================================================
                            Azure Custom Image Builder Script
===============================================================================
Creator      : Marc-André Brochu
Email        : mabrochu@it-ed.com
Company      : ITED
Date         : 14 Novembre 2024
Description  : This script is used with Azure Custom Image Builder to:
               - Download Office setup.exe and configuration XML files.
               - Log installation steps with a specific prefix (AIB-).
               - Install Office using the downloaded files with the required steps.
===============================================================================
#>

# Variables
$OfficeSetupUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-Office/setup.exe"
$OfficeConfigUrl = "https://raw.githubusercontent.com/it-ed-com/RMM-Public/refs/heads/main/Custom%20image%20templates/Install-Office/configuration-Office365-x64.xml"
$OfficePath = "C:\ited\office365"
$LogPath = "C:\ited\log"
$LogPrefix = "AIB-Office365"
$OfficeSetupExe = Join-Path -Path $OfficePath -ChildPath "setup.exe"
$OfficeConfigXml = Join-Path -Path $OfficePath -ChildPath "configuration-Office365-x64.xml"
$SourcePath = Join-Path -Path $OfficePath -ChildPath "OfficeSources"

# Ensure folders exist
New-Item -ItemType Directory -Path $OfficePath -Force | Out-Null
New-Item -ItemType Directory -Path $LogPath -Force | Out-Null

# Start transcript for logging
$LogFile = Join-Path -Path $LogPath -ChildPath "$LogPrefix-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
Start-Transcript -Path $LogFile -NoClobber

try {
    Write-Host "Downloading Office setup.exe..."
    Invoke-WebRequest -Uri $OfficeSetupUrl -OutFile $OfficeSetupExe -ErrorAction Stop

    Write-Host "Downloading Office configuration XML..."
    Invoke-WebRequest -Uri $OfficeConfigUrl -OutFile $OfficeConfigXml -ErrorAction Stop

    # Download Office sources
    Write-Host "Downloading Office sources..."
    Start-Process -FilePath $OfficeSetupExe -ArgumentList "/download `"$OfficeConfigXml`"" -Wait -NoNewWindow -ErrorAction Stop

    # Install Office
    Write-Host "Installing Office..."
    Start-Process -FilePath $OfficeSetupExe -ArgumentList "/configure `"$OfficeConfigXml`"" -Wait -NoNewWindow -ErrorAction Stop

    Write-Host "Office installation completed successfully."
} catch {
    Write-Error "An error occurred: $_"
} finally {
    # Stop transcript
    Stop-Transcript
}
