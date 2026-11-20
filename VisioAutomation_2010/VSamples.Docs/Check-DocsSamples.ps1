<#
.SYNOPSIS
    Checks that the C# code blocks in the .NET GitBook docs match the samples in Samples\Docs.

.DESCRIPTION
    Every C# block in the docs has an ID: the page path without .md, '#', and the block's
    position among the C# blocks on that page (for example models/directed-graph#03).
    Samples\Docs holds one method per block, with the block's text between
    '// <sample id="...">' and '// </sample>' markers.

    The check reports:
      * a docs block with no sample (and no entry in Samples\Docs\exempt.tsv),
      * a docs block whose text differs from its sample,
      * a sample whose block is no longer in the docs.

    Lines that are 'using' directives are not compared: the sample files keep them at the top
    of the file. Run Build first; this script does not compile anything.

.PARAMETER DocsPath
    Path to a checkout of VisioAutomation_GitBook_Docs.

.EXAMPLE
    .\Check-DocsSamples.ps1 -DocsPath C:\src\VisioAutomation_GitBook_Docs
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $DocsPath
)

$ErrorActionPreference = 'Stop'

$samplesRoot = Join-Path $PSScriptRoot 'Samples\Docs'
$usingLine = '^\s*using\s+(static\s+)?[\w.]+(\s*=\s*[\w.]+)?\s*;\s*(//.*)?$'

function Get-Lines([string] $text) { return ,($text -split "\r?\n") }

function Format-Block([string[]] $lines) {
    # Drop trailing spaces, then leading and trailing blank lines.
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($l in $lines) { $out.Add($l.TrimEnd()) }
    while ($out.Count -gt 0 -and $out[0] -eq '') { $out.RemoveAt(0) }
    while ($out.Count -gt 0 -and $out[$out.Count - 1] -eq '') { $out.RemoveAt($out.Count - 1) }
    return ($out -join "`n")
}

# --- exempt ids
$exempt = @{}
$exemptFile = Join-Path $samplesRoot 'exempt.tsv'
if (Test-Path $exemptFile) {
    foreach ($row in (Get-Content $exemptFile)) {
        $parts = $row -split "`t"
        if ($parts.Count -ge 4) { $exempt[$parts[0]] = $parts[3] }
    }
}

# --- samples: id -> text between the markers
$samples = @{}
$markerPattern = '(?s)//\s*<sample id="([^"]+)"[^>]*>\r?\n(.*?)\r?\n\s*//\s*</sample>'
foreach ($file in (Get-ChildItem -Path $samplesRoot -Recurse -Filter '*.cs')) {
    $text = [System.IO.File]::ReadAllText($file.FullName)
    foreach ($m in [regex]::Matches($text, $markerPattern)) {
        $body = Get-Lines $m.Groups[2].Value
        # The generated methods indent the block by 12 spaces.
        $dedented = foreach ($l in $body) { if ($l.StartsWith('            ')) { $l.Substring(12) } else { $l.TrimStart() } }
        $samples[$m.Groups[1].Value] = Format-Block @($dedented)
    }
}

# --- docs blocks: id -> text
$docsRoot = (Resolve-Path $DocsPath).Path
$blocks = @{}
foreach ($md in (Get-ChildItem -Path $docsRoot -Recurse -Filter '*.md' | Where-Object { $_.FullName -notmatch '[\\/]\.git[\\/]' })) {
    $rel = $md.FullName.Substring($docsRoot.Length).TrimStart('\', '/') -replace '\\', '/'
    $stem = $rel.Substring(0, $rel.Length - 3)
    $lines = Get-Lines ([System.IO.File]::ReadAllText($md.FullName))
    $n = 0
    $i = 0
    while ($i -lt $lines.Count) {
        if ($lines[$i] -match '^\s*```csharp\s*$') {
            $n++
            $body = New-Object System.Collections.Generic.List[string]
            $i++
            while ($i -lt $lines.Count -and $lines[$i] -notmatch '^\s*```\s*$') {
                if ($lines[$i] -notmatch $usingLine) { $body.Add($lines[$i]) }
                $i++
            }
            $blocks[('{0}#{1:D2}' -f $stem, $n)] = Format-Block $body.ToArray()
        }
        $i++
    }
}

# --- compare
$missing = @(); $different = @(); $orphans = @(); $exempted = 0
foreach ($id in ($blocks.Keys | Sort-Object)) {
    if ($exempt.ContainsKey($id)) { $exempted++; continue }
    if (-not $samples.ContainsKey($id)) { $missing += $id; continue }
    if ($samples[$id] -ne $blocks[$id]) { $different += $id }
}
foreach ($id in ($samples.Keys | Sort-Object)) {
    if (-not $blocks.ContainsKey($id)) { $orphans += $id }
}

Write-Host ("Docs blocks: {0}; samples: {1}; exempt: {2}" -f $blocks.Count, $samples.Count, $exempted)
if ($missing.Count)   { Write-Host "No sample for these docs blocks (add one, or list it in exempt.tsv):"; $missing   | ForEach-Object { Write-Host "  $_" } }
if ($different.Count) { Write-Host "The docs block and the sample differ:";                              $different | ForEach-Object { Write-Host "  $_" } }
if ($orphans.Count)   { Write-Host "Samples whose docs block is gone:";                                  $orphans   | ForEach-Object { Write-Host "  $_" } }

if ($missing.Count -or $different.Count -or $orphans.Count) { exit 1 }
Write-Host 'OK: every docs block has a matching sample.'
