# PURPOSE
# -------
# Downloads the latest published Visio PowerShell module from the PowerShell Gallery
# (https://www.powershellgallery.com/packages/Visio) and imports it into the current
# session.
#
# USE CASE
# --------
# Verify that the gallery-published module actually loads and works in a clean session,
# independent of the local bin/Debug build. Useful as a release-verification step
# before or after publishing a new version.
#
# REQUIREMENTS
# ------------
# - PowerShell 5.0 or later
# - PowerShellGet (built in to PS 5.1)
# - Internet access to www.powershellgallery.com
# - PSGallery registered as a trusted repository, or accept the prompt at runtime

Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

# Save the module to a local subfolder (gitignored), not the user's PS modules folder.
# This keeps the download isolated and easy to inspect or delete.
$download_folder = Join-Path $PSScriptRoot "DownloadedModule"

if (-not (Test-Path $download_folder))
{
    New-Item -ItemType Directory -Path $download_folder | Out-Null
}

Write-Host "Downloading 'Visio' module from PowerShell Gallery to:" $download_folder
Save-Module -Name Visio -Repository PSGallery -Path $download_folder -Force

# Save-Module places the module at <download_folder>/Visio/<version>/Visio.psd1
$psd1 = Get-ChildItem -Path (Join-Path $download_folder "Visio") -Recurse -Filter "Visio.psd1" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $psd1)
{
    Write-Error "Could not locate Visio.psd1 under $download_folder"
}

Write-Host "Importing module from:" $psd1.FullName
Import-Module $psd1.FullName -Verbose
