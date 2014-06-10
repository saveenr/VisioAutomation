#NOTE: Please remove any commented lines to tidy up prior to releasing the package, including this one


$packageName = 'Visio PowerShell Module' # arbitrary name for the package, used in messages
$installerType = 'MSI' #only one of these: exe, msi, msu
$url = "VisioPS_1.2.62.msi" # download url
$silentArgs = '/quiet' 
$validExitCodes = @(0) 


Install-ChocolateyPackage "$packageName" "$installerType" "$silentArgs" "$url" -validExitCodes $validExitCodes

