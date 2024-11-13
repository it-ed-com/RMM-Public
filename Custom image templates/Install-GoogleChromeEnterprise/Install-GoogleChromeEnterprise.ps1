<#
===============================================================================
                            Azure Custom Image Builder Script
===============================================================================
Creator      : Marc-Andre Brochu
Email        : mabrochu@it-ed.com
Date         : 12 Novembre 2024
Description  : This script installs Adobe Reader using winget on a custom Azure image.
               The installation process logs all activities into C:\ited\logs with
               a specific prefix for easy identification. Attempts to install the
               French version if available.
===============================================================================
#>

# Define paths for logs and temporary files
$logDir = "C:\ited\logs"
$tempDir = "C:\ited\temp"
$logFile = "$logDir\aib-adobe_reader_install.log"

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

    # Install Adobe Reader for all users using winget
    Write-Output "Starting Adobe Reader installation..."
    # Attempt to install the French version of Adobe Reader
    winget install --id Adobe.Acrobat.Reader.32-bit --scope machine --silent --accept-package-agreements --accept-source-agreements --locale fr-FR -ErrorAction SilentlyContinue
    if ($LASTEXITCODE -eq 0) {
        Write-Output "French version of Adobe Reader installed successfully."
    } else {
        Write-Output "French version not available, installing default version of Adobe Reader..."
        winget install --id Adobe.Acrobat.Reader.32-bit --scope machine --silent --accept-package-agreements --accept-source-agreements
    }
    Write-Output "Adobe Reader installation completed successfully."
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
