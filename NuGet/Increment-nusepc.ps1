param
(
	[parameter(Mandatory=$true)] [string] $File,
	[parameter(Mandatory=$false)] [Switch] $Major,
	[parameter(Mandatory=$false)] [Switch] $Minor,
	[parameter(Mandatory=$false)] [Switch] $Patch
)
Set-StrictMode -Version 2 
$ErrorActionPreference = "Stop"


$File = resolve-path $File
$Extension = [System.IO.Path]::GetExtension($File)

if ($Extension -ne ".nuspec")
{
	Write-Host This Script only works on NuSpec files
	break
}

$xml = [xml] (Get-Content $File)

$OldVersion = $xml.package.metadata.version
$tokens = $OldVersion.Split(".")
Write-Host Old Version = $OldVersion

$OldMajor = [int] $tokens[0]
$OldMinor = [int] $tokens[1]
$OldPatch = [int] $tokens[2]

if ( !$Major -and !$Minor -and !$Patch)
{
	Write-Host No input provided, will increment patch
	$Patch = $true
}

if ($Major)
{
	$NewMajor = $OldMajor + 1
	$NewMinor = 0
	$NewPatch = 0
}
elseif ($Minor)
{
	$NewMajor = $OldMajor
	$NewMinor = $OldMinor + 1
	$NewPatch = 0
}
elseif ($Patch)
{
	$NewMajor = $OldMajor
	$NewMinor = $OldMinor
	$NewPatch = $OldPatch +1
}

$NewVersion = "$NewMajor.$NewMinor.$NewPatch"

Write-Host New Version = $NewVersion

$xml.package.metadata.version = $NewVersion

$xml.Save( $File )

