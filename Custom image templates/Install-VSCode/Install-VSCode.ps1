<#
===============================================================================
                            Azure Custom Image Builder Script
===============================================================================
Creator      : Marc-Andre Brochu
Email        : mabrochu@it-ed.com
Date         : 12 Novembre 2024
Description  : This script installs Visual Studio Code using winget on a custom Azure image.
               The installation process logs all activities into C:\ited\logs with
               a specific prefix for easy identification.
===============================================================================
#>

# Define paths for logs and temporary files
$logDir = "C:\ited\logs"
$tempDir = "C:\ited\temp"
$logFile = "$logDir\aib-vscode_install.log"

# Create directories if they do not exist
if (-not (Test-Path -Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir -Force
}

if (-not (Test-Path -Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force
}

# Start logging
Start-Transcript -Path $logFile -Append

try {
    # Verify if winget is installed
    if (-not (Get-Command -Name "winget" -ErrorAction SilentlyContinue)) {
        $errorMessage = "winget not found. Ensure Windows Package Manager is installed."
        Write-Output $errorMessage
        throw $errorMessage
    }

    # Install Visual Studio Code for all users using winget
    Write-Output "Starting Visual Studio Code installation..."
    winget install --id Microsoft.VisualStudioCode --scope machine --silent --accept-package-agreements --accept-source-agreements
    Write-Output "Visual Studio Code installation completed successfully."
}
catch {
    # Log any errors that occur
    $errorMessage = "An error occurred: $_"
    Write-Output $errorMessage
}
finally {
    # End logging
    Stop-Transcript
}

Write-Output "Script execution completed."
