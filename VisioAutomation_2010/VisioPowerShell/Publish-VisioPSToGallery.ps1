# PURPOSE
# -------
# Single-command release of the Visio PowerShell module to the PowerShell Gallery.
#
# Steps performed:
#   1. Read the version from Visio.psd1
#   2. Stage the Debug build into the user's modules folder (calls InstallForCurrentUser.ps1)
#   3. Publish the staged module to the PowerShell Gallery
#   4. Tag HEAD as VisioPS_<version> and push the tag to origin
#
# USAGE
# -----
#     .\Publish-VisioPSToGallery.ps1 -ApiKey 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
#
# Or set the API key via env var and run with no arguments:
#     $env:PSGalleryApiKey = 'xxxx...'
#     .\Publish-VisioPSToGallery.ps1
#
# PREREQUISITES
# -------------
# - You must be the package owner of "Visio" on the PowerShell Gallery.
# - The solution must already be built (Debug). Run MSBuild beforehand if needed.
# - Working tree should be clean and HEAD should be the commit you want to tag.
# - Use a fresh PowerShell session — if another session has the Visio module
#   loaded, the staging step will fail (locked DLLs).

[CmdletBinding()]
param(
    [string]$ApiKey,

    # Skip the publish step (useful for testing the staging + tag workflow without
    # actually pushing to the gallery).
    [switch]$WhatIf
)

Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

# ----- Resolve API key -----
if (-not $ApiKey) {
    if ($env:PSGalleryApiKey) {
        $ApiKey = $env:PSGalleryApiKey
    }
    elseif (-not $WhatIf) {
        throw "API key not provided. Pass -ApiKey or set `$env:PSGalleryApiKey."
    }
}

# ----- Paths -----
$script_path   = $MyInvocation.MyCommand.Path
$script_folder = Split-Path $script_path -Parent
$psd1_path     = Join-Path $script_folder 'Visio.psd1'
$install_script = Join-Path $script_folder 'InstallForCurrentUser.ps1'

if (-not (Test-Path $psd1_path)) {
    throw "Visio.psd1 not found at $psd1_path"
}
if (-not (Test-Path $install_script)) {
    throw "InstallForCurrentUser.ps1 not found at $install_script"
}

# ----- Read version from manifest -----
$manifest = Import-PowerShellDataFile $psd1_path
$version  = $manifest.ModuleVersion
if (-not $version) {
    throw "Could not read ModuleVersion from $psd1_path"
}
$tag = "VisioPS_$version"

Write-Host ""
Write-Host "===================================================="
Write-Host "Publishing Visio PowerShell module to PSGallery"
Write-Host "----------------------------------------------------"
Write-Host "Version : $version"
Write-Host "Tag     : $tag"
Write-Host "WhatIf  : $WhatIf"
Write-Host "===================================================="
Write-Host ""

# ----- Step 1: stage the build -----
Write-Host "[1/3] Staging module via InstallForCurrentUser.ps1 ..."
& $install_script
if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) {
    throw "InstallForCurrentUser.ps1 exited with code $LASTEXITCODE"
}

# ----- Step 2: publish -----
if ($WhatIf) {
    Write-Host "[2/3] WhatIf: skipping Publish-Module."
}
else {
    Write-Host "[2/3] Publishing to PowerShell Gallery ..."
    Publish-Module -Name "Visio" -NuGetApiKey $ApiKey
}

# ----- Step 3: tag and push -----
$repo_root = Resolve-Path (Join-Path $script_folder '..\..')
Push-Location $repo_root
try {
    # Refuse to tag if working tree is dirty.
    $status = git status --porcelain
    if ($status) {
        throw "Working tree is not clean. Commit or stash before tagging."
    }

    # Refuse to tag if HEAD doesn't match origin's tracking branch (avoid tagging
    # a commit nobody else can fetch).
    $branch = git rev-parse --abbrev-ref HEAD
    $local_head  = git rev-parse HEAD
    $upstream    = git rev-parse "@{u}" 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $upstream) {
        throw "Branch '$branch' has no upstream — push it before tagging."
    }
    if ($local_head -ne $upstream) {
        throw "Local '$branch' is not in sync with origin. Push commits before tagging."
    }

    # Refuse to overwrite an existing tag.
    $existing = git tag -l $tag
    if ($existing) {
        throw "Tag '$tag' already exists. Bump the version in Visio.psd1 if you need to re-release."
    }

    if ($WhatIf) {
        Write-Host "[3/3] WhatIf: skipping git tag and push."
    }
    else {
        Write-Host "[3/3] Tagging $tag and pushing ..."
        git tag $tag
        git push origin $tag
    }
}
finally {
    Pop-Location
}

Write-Host ""
Write-Host "Done."
Write-Host "  Module : Visio $version"
Write-Host "  Gallery: https://www.powershellgallery.com/packages/Visio/$version"
Write-Host "  Tag    : $tag"
