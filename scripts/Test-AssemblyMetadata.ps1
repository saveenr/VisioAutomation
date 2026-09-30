param(
    [Parameter(Mandatory)][string] $Solution,
    [Parameter(Mandatory)][string] $PackageMetadata,
    [string] $ModuleManifest,
    [ValidateSet('Debug', 'Release')][string] $Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$solutionPath = Join-Path $root $Solution
[xml] $solutionXml = Get-Content -Raw -Encoding UTF8 -LiteralPath $solutionPath
[xml] $package = Get-Content -Raw -Encoding UTF8 -LiteralPath (Join-Path $root $PackageMetadata)
$author = 'VisioAutomation contributors'
if ($package.package.metadata.authors -ne $author) { throw 'Unexpected package authorship.' }
if ($package.package.metadata.owners) { throw 'Remove obsolete nuspec owners metadata.' }

$module = $null
if ($ModuleManifest) {
    $module = Import-PowerShellDataFile (Join-Path $root $ModuleManifest)
    if ($module.Author -ne $author) { throw 'Unexpected module authorship.' }
}
$revision = $null
if (Test-Path -LiteralPath (Join-Path $root '.git')) {
    $revision = (& git -C $root rev-parse HEAD).Trim()
    if ($LASTEXITCODE -ne 0) { throw 'Cannot determine the source revision.' }
}

function Get-AttributeValue($Attributes, [string] $Name) {
    $matching = @($Attributes | Where-Object { $_.AttributeType.FullName -eq $Name })
    if ($matching.Count -ne 1) { throw "Expected exactly one $Name attribute." }
    return $matching[0].ConstructorArguments[0].Value
}

$count = 0
$projects = $solutionXml.SelectNodes('//Project')
if ($projects.Count -eq 0) { throw 'The solution contains no projects to validate.' }
foreach ($entry in $projects) {
    $projectPath = Join-Path (Split-Path $solutionPath -Parent) $entry.Path
    [xml] $project = Get-Content -Raw -Encoding UTF8 -LiteralPath $projectPath
    $nameNode = $project.SelectSingleNode('//AssemblyName')
    $name = if ($nameNode) { $nameNode.InnerText } else { [IO.Path]::GetFileNameWithoutExtension($projectPath) }
    $outputNode = $project.SelectSingleNode('//OutputType')
    $extension = if ($outputNode -and $outputNode.InnerText -in @('Exe', 'WinExe')) { 'exe' } else { 'dll' }
    $path = Join-Path (Split-Path $projectPath -Parent) "bin\$Configuration\$name.$extension"
    $release = [string] $package.package.metadata.version
    if ($module -and $name -eq [IO.Path]::GetFileNameWithoutExtension($module.RootModule)) {
        $release = [string] $module.ModuleVersion
    }
    $numeric = [Version] ($release.Split([char[]] '-+')[0])
    $fileVersion = '{0}.{1}.{2}.{3}' -f $numeric.Major, $numeric.Minor, [Math]::Max(0, $numeric.Build), [Math]::Max(0, $numeric.Revision)

    # A fresh PowerShell process avoids inspecting an older, already-loaded build.
    $assembly = [Reflection.Assembly]::LoadFrom($path)
    $identity = $assembly.GetName()
    if ($identity.Name -ne $name -or $identity.Version.ToString() -ne '1.0.0.0' -or $identity.GetPublicKeyToken().Length -ne 0) {
        throw "Unexpected assembly identity: $path"
    }
    $attributes = [Reflection.CustomAttributeData]::GetCustomAttributes($assembly)
    $actualFileVersion = Get-AttributeValue $attributes 'System.Reflection.AssemblyFileVersionAttribute'
    if ($actualFileVersion -ne $fileVersion) { throw "$name file version $actualFileVersion differs from release $fileVersion." }
    $information = Get-AttributeValue $attributes 'System.Reflection.AssemblyInformationalVersionAttribute'
    $expectedInformation = if ($revision) { "$release+$revision" } else { $release }
    if ($information -ne $expectedInformation) { throw "$name informational version $information differs from $expectedInformation." }
    if ((Get-AttributeValue $attributes 'System.Reflection.AssemblyConfigurationAttribute') -ne $Configuration) {
        throw "$name has the wrong build configuration."
    }
    if ((Get-AttributeValue $attributes 'System.Reflection.AssemblyCompanyAttribute') -ne $author) {
        throw "$name has unexpected contributor attribution."
    }
    foreach ($attribute in @('AssemblyTitle', 'AssemblyProduct', 'AssemblyDescription', 'AssemblyCopyright')) {
        if ([string]::IsNullOrWhiteSpace((Get-AttributeValue $attributes "System.Reflection.${attribute}Attribute"))) {
            throw "$name is missing $attribute."
        }
    }
    if ((Get-AttributeValue $attributes 'System.Runtime.InteropServices.ComVisibleAttribute') -ne $false) {
        throw "$name must not expose COM types."
    }
    $count++
    Write-Host "${name}: identity 1.0.0.0; file $actualFileVersion; informational $information"
}
Write-Host "Metadata checks passed for $count $Configuration assemblies."
