Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

$sourceNugetExe = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
$destination_path = [Environment]::GetFolderPath("MyDocuments")
$targetNugetExe = "$destination_path\nuget.exe"

Write-Host "Downloading install NuGet to $destination_path"

if (!(test-path $destination_path))
{
    New-Item -Path $destination_path -ItemType Directory
}


if (test-path $targetNugetExe)
{
    Remove-Item $targetNugetExe
}

Invoke-WebRequest $sourceNugetExe -OutFile $targetNugetExe
Set-Alias nuget $targetNugetExe -Scope Global -Verbose