./pubinfo.ps1
$url="https://www.nuget.org"

Write-Host Publishing $global:packagefilename 
.\nuget.exe push $global:packagefilename -Source $url