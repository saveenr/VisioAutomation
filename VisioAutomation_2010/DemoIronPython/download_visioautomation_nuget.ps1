


function makedirsafe ($p)
{
    if (-not (Test-Path($p)))
    {
        $result = New-Item $p -type directory
    }

}

function download_nuget_package( $packagename  )
{ 
    $url = "http://www.nuget.org/api/v2/package/" + $packagename
    $scriptpath = $PSScriptRoot

    $pkgs_folder = Join-Path $scriptpath "packages"
    $pkgs_va_folder = Join-Path $pkgs_folder  $packagename 


    makedirsafe $pkgs_folder

    if (Test-Path $pkgs_va_folder)
    {
        Remove-Item -Recurse -Force $pkgs_va_folder
    }
 
    makedirsafe $pkgs_va_folder 

    $zipfile = Join-Path $pkgs_folder  ( $packagename  + ".zip" )
    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile($url, $zipfile)


    $asm = [System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $pkgs_va_folder)
    Remove-Item $zipfile

}


download_nuget_package "VisioAutomation2010"