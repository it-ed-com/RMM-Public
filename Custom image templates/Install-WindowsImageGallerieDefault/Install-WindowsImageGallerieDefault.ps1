<#
    Script Name: Install-PhotoViewer.ps1
    Author:      Marc-Andre Brochu
    Email:       mabrochu@it-ed.com
    Date:        October 17, 2024
    Description: This script re-enables Microsoft Photo Viewer and sets it as the default application
                 for .bmp, .jpe, .jpeg, and .jpg file extensions on Windows 10 Multi-Enterprise.
                 It logs all actions and errors to C:\ited\logs\photoviewer.log.
#>

# =========================================
# Configuration Parameters
# =========================================

# Ensure the log directory exists
$logDirectory = "C:\ited\logs"
if (!(Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
}

# Log file path
$logFile = "$logDirectory\photoviewer.log"

# File extensions to associate with Photo Viewer
$fileExtensions = @(".bmp", ".jpe", ".jpeg", ".jpg")

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
# Validation Function
# =========================================

Function Is-PhotoViewerEnabled {
    # Check if Photo Viewer is enabled in registry
    $regPath = "HKLM:\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations"
    if (Test-Path -Path $regPath) {
        return $true
    }
    else {
        return $false
    }
}

# Function to enable Photo Viewer in registry
Function Enable-PhotoViewer {
    Write-Log "Enabling Microsoft Photo Viewer in registry."

    # Define the registry keys and values
    $regKeys = @{
        "HKLM:\SOFTWARE\Microsoft\Windows Photo Viewer\Capabilities\FileAssociations" = @{
            ".bmp" = "PhotoViewer.FileAssoc.Bmp"
            ".jpe" = "PhotoViewer.FileAssoc.Jpeg"
            ".jpeg" = "PhotoViewer.FileAssoc.Jpeg"
            ".jpg" = "PhotoViewer.FileAssoc.Jpeg"
        }
    }

    foreach ($path in $regKeys.Keys) {
        foreach ($ext in $regKeys[$path].Keys) {
            $progId = $regKeys[$path][$ext]
            # Create or set the registry key
            New-Item -Path $path -Force | Out-Null
            Set-ItemProperty -Path $path -Name $ext -Value $progId -Force
            Write-Log "Set $ext to $progId in $path."
        }
    }

    Write-Log "Microsoft Photo Viewer enabled in registry."
}

# Function to set file association using assoc and ftype
Function Set-FileAssociation {
    param (
        [string]$extension,
        [string]$progId
    )

    Write-Log "Associating $extension with $progId."

    # Use assoc to associate the extension with the ProgId
    $assocCmd = "assoc $extension=$progId"
    Write-Log "Executing: $assocCmd"
    cmd.exe /c $assocCmd

    # Define the ftype for the ProgId
    $ftypeCmd = "ftype $progId=`"rundll32.exe %SystemRoot%\System32\shimgvw.dll,ImageView_Fullscreen %1`""
    Write-Log "Executing: $ftypeCmd"
    cmd.exe /c $ftypeCmd
}

# =========================================
# Script Execution
# =========================================

Try {
    Write-Log "Starting Microsoft Photo Viewer installation and association script."

    # Ensure script is running as Administrator
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole] "Administrator")) {
        Write-Log "Script is not running with administrative privileges." "ERROR"
        Throw "Please run this script as an Administrator."
    }

    # Enable Photo Viewer in registry if not already enabled
    if (-not (Is-PhotoViewerEnabled)) {
        Enable-PhotoViewer
    }
    else {
        Write-Log "Microsoft Photo Viewer is already enabled in registry."
    }

    # Set file associations
    foreach ($ext in $fileExtensions) {
        switch ($ext) {
            ".bmp" { $progId = "PhotoViewer.FileAssoc.Bmp" }
            ".jpe" { $progId = "PhotoViewer.FileAssoc.Jpeg" }
            ".jpeg" { $progId = "PhotoViewer.FileAssoc.Jpeg" }
            ".jpg" { $progId = "PhotoViewer.FileAssoc.Jpeg" }
            default { $progId = "PhotoViewer.FileAssoc.Jpeg" }
        }

        Set-FileAssociation -extension $ext -progId $progId
    }

    Write-Log "File associations set successfully."

    Write-Log "Microsoft Photo Viewer installation and association script completed successfully."
    Exit 0
}
Catch {
    Write-Log "An error occurred: $_" "ERROR"
    Exit 1
}
Finally {
    # Nothing needed here
}
