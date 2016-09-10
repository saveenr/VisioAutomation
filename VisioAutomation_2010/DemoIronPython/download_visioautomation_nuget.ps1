$scriptpath = $PSScriptRoot

$pkgs_folder = Join-Path $scriptpath "packages"
$pkgs_va_folder = Join-Path $pkgs_folder  "VisioAutomation"


function makedirsafe ($p)
{
    if (-not (Test-Path($p)))
    {
        New-Item $p -type directory
    }

}

makedirsafe $pkgs_folder 
makedirsafe $pkgs_va_folder 

$url = "http://www.nuget.org/api/v2/package/VisioAutomation2010"
$zipfile = Join-Path $pkgs_folder  "VisioAutomation2010.zip"
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($url, $zipfile)


[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')
[System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $pkgs_va_folder)
