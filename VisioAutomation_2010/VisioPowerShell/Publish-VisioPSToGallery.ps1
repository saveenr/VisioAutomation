# PURPOSE
# -------
# Single-command release of the Visio PowerShell module to the PowerShell Gallery.
#
# Steps performed:
#   1. Read the version from Visio.psd1
#   2. Stage the Release build into the user's modules folder (calls
#      InstallForCurrentUser.ps1 -Configuration Release)
#   3. Publish the staged module to the PowerShell Gallery
#   4. Tag HEAD as VisioPS_<version> and push the tag to origin (idempotent:
#      if the tag already exists at HEAD, skipped silently -- supports running
#      this script after .github/workflows/publish-psmodule.yml has already
#      created the tag).
#
# CANONICAL RELEASE FLOW
# ----------------------
# The preferred release flow is now CI-driven:
#   1. Trigger .github/workflows/release-psmodule.yml manually. Builds Release,
#      stages the module, creates the GitHub Release tagged VisioPS_<version>
#      with the staged module zip attached.
#   2. Trigger .github/workflows/publish-psmodule.yml with the tag from step 1.
#      Downloads the GH Release zip, publishes to PSGallery, verifies via
#      Find-Module.
# This script remains as a fallback / dev-convenience path for out-of-band
# publishes. It coexists with the workflow because the tag step here is
# idempotent.
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
# - The solution must already be built in Release. Run
#   `msbuild VisioAutomation_2010\VisioAutomation2010.sln -p:Configuration=Release -m`
#   beforehand if needed.
# - Working tree should be clean and HEAD should be the commit you want to tag.
# - Use a fresh PowerShell session -- if another session has the Visio module
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

# ----- Force TLS 1.2 -----
# PS 5.1 defaults to TLS 1.0/1.1 in many configs; PSGallery requires 1.2.
# Without this, Publish-Module fails with 'Could not create SSL/TLS secure
# channel'. PS 7 negotiates TLS 1.2 automatically, but doing it here keeps
# the script behavior identical across editions.
[Net.ServicePointManager]::SecurityProtocol =
    [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# ----- Explicit PowerShell host check -----
# Different PowerShell editions use different user-module paths and have
# different Publish-Module versions. Capture what we're running so it's
# visible in the banner below, and refuse to proceed on anything below 5.1.
$ps_version = $PSVersionTable.PSVersion
$ps_edition = if ($PSVersionTable.PSObject.Properties.Name -contains 'PSEdition') {
    $PSVersionTable.PSEdition
}
else {
    'Desktop'   # PS 5.1 ships with PSEdition; PS 2.0-3.0 don't.
}
$ps_min = [Version]'5.1'
if ($ps_version -lt $ps_min) {
    throw "PowerShell $ps_version is too old to publish. Need at least $ps_min (Windows PowerShell 5.1) or PowerShell 7+."
}

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
Write-Host "Host    : PowerShell $ps_version ($ps_edition), PID $PID"
Write-Host "WhatIf  : $WhatIf"
Write-Host "===================================================="
Write-Host ""
if ($ps_edition -eq 'Core') {
    Write-Host "Note: running under PowerShell 7+. Using Publish-Module -Path"
    Write-Host "      to bypass the PS 5.1 vs 7 user-module path split."
    Write-Host ""
}

# ----- Step 1: stage the build -----
Write-Host "[1/3] Staging Release build via InstallForCurrentUser.ps1 -Configuration Release ..."
& $install_script -Configuration Release
# Don't trust $LASTEXITCODE here -- InstallForCurrentUser.ps1's last native
# command is robocopy, which uses bit-flag exit codes where 0 means "no files
# copied" and 1 means "files copied successfully". Both are non-failures, but
# anything that propagates as $LASTEXITCODE looks like an error to a naive
# check. Verify the actual outcome below instead.

# Verify the staged module folder has the expected version.
$staged_psd1 = Join-Path $home 'Documents\WindowsPowerShell\Modules\Visio\Visio.psd1'
if (-not (Test-Path $staged_psd1)) {
    throw "Staged module not found at $staged_psd1 -- InstallForCurrentUser.ps1 may have failed."
}
$staged_version = (Import-PowerShellDataFile $staged_psd1).ModuleVersion
if ($staged_version -ne $version) {
    throw "Staged module is version $staged_version but the source manifest is $version. Rebuild the solution (MSBuild Debug) so the bin/Debug copy of Visio.psd1 reflects the bumped version, then re-run."
}
$staged_folder = Split-Path $staged_psd1 -Parent
Write-Host "      Staged Visio $staged_version OK at $staged_folder"

# ----- Step 2: publish -----
# Use -Path (not -Name) so Publish-Module locates the module deterministically
# regardless of the running shell's PSModulePath. PowerShell 7 and Windows
# PowerShell 5.1 use different user-module folders, and the staging script
# writes to the 5.1 path; -Path sidesteps the discovery mismatch entirely.
#
# -ErrorAction Stop is necessary because PowerShellGet 1.x (the in-box version
# on Windows PowerShell 5.1) catches its own internal failures and re-emits
# them via Write-Error in a nested scope -- without -ErrorAction Stop on the
# outer call, the script-level $ErrorActionPreference doesn't reliably stop
# the script. Even with that, we verify the publish positively below by
# querying PSGallery; do not skip that check.
if ($WhatIf) {
    Write-Host "[2/3] WhatIf: skipping Publish-Module."
}
else {
    Write-Host "[2/3] Publishing to PowerShell Gallery ..."
    Publish-Module -Path $staged_folder -NuGetApiKey $ApiKey -ErrorAction Stop

    # Positive verification: query PSGallery for the version we just published.
    # Catches the case where Publish-Module emitted Write-Error but the script
    # didn't terminate (PowerShellGet 1.x bug). Retry briefly because the
    # gallery's version index is sometimes a few seconds behind the upload.
    Write-Host "      Verifying $version is live on PSGallery ..."
    $found = $null
    for ($attempt = 1; $attempt -le 6; $attempt++) {
        Start-Sleep -Seconds 5
        try {
            $found = Find-Module -Name 'Visio' -RequiredVersion $version `
                -Repository 'PSGallery' -ErrorAction Stop
            if ($found) { break }
        }
        catch {
            # Module not yet listed -- keep retrying until timeout
            if ($attempt -eq 6) { throw }
        }
    }
    if (-not $found) {
        throw "Publish appeared to run, but version $version is not visible on PSGallery after 30s. Check the gallery directly: https://www.powershellgallery.com/packages/Visio/$version"
    }
    Write-Host "      Confirmed: Visio $version is on PSGallery."
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
        throw "Branch '$branch' has no upstream -- push it before tagging."
    }
    if ($local_head -ne $upstream) {
        throw "Local '$branch' is not in sync with origin. Push commits before tagging."
    }

    # Refresh local tag knowledge so we see tags that may have been created
    # by the publish-psmodule.yml CI workflow but not yet fetched locally.
    git fetch --tags origin --quiet 2>$null

    # Tag-creation rules (idempotent so this script can run after the CI
    # workflow has already tagged):
    #   - Tag does not exist:                    create + push
    #   - Tag exists and points to HEAD:         re-push (no-op if already on origin)
    #   - Tag exists and points elsewhere:       fail loudly (refuse to silently re-tag)
    $existing = git tag -l $tag
    if ($existing) {
        $tag_sha  = git rev-parse "$tag^{commit}"
        $head_sha = git rev-parse HEAD
        if ($tag_sha -ne $head_sha) {
            throw "Tag '$tag' already exists at $tag_sha, but HEAD is $head_sha. Refusing to silently re-tag. Bump the version in Visio.psd1 if you need to re-release."
        }
    }

    if ($WhatIf) {
        Write-Host "[3/3] WhatIf: skipping git tag and push."
    }
    elseif ($existing) {
        Write-Host "[3/3] Tag '$tag' already exists at HEAD; ensuring it's pushed to origin."
        # `git push origin <tag>` is a no-op for tags already on origin.
        git push origin $tag
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
