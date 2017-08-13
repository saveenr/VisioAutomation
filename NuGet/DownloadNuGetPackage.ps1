Set-StrictMode -Version 2
$ErrorActionPreference = "Stop"

$package_name = "VisioAutomation2010"
$pkgsource_name = "nuget.org"
$destination_path = [Environment]::GetFolderPath("MyDocuments")

Write-Host "Attempting to install NuGet $package_name to $destination_path"

$pkgsource = Get-PackageSource -Name $pkgsource_name 
$package = find-package $package_name -Source $pkgsource.Location

if (!(test-path $destination_path))
{
    New-Item -Path $destination_path -ItemType Directory
}

$package | Install-Package -Destination $destination_path -Force


