



$modulename = "VisioPS"

Write-Host Installing $modulename module


$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath

$src_path = $dir
write-host "FROM" $dir

$mydocs = [Environment]::GetFolderPath("MyDocuments")
$psmodfolder = join-path $mydocs ( "WindowsPowerShell\Modules\" + $modulename )

write-host "TO" $psmodfolder

if (!(test-path $psmodfolder) )
{
    New-Item $psmodfolder -type directory
}

robocopy $src_path $psmodfolder /mir





