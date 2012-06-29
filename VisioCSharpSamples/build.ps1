param([string]$param_pub)
$pub_params = $param_pub -split ","

$localfeed ="D:\saveenr\live-mesh\code\mynugetfeed" 
$publicfeedurl="https://www.nuget.org"
$packaging_folder = ".\Packaging" 

$publocal = $pub_params.Contains( "local" )
$pubremote = $pub_params.Contains( "remote" )


$nuspecfilname = "Package.nuspec"

Write-Host looading $nuspecfilname 

if (!(test-path $nuspecfilname))
{
    Write-Host ERROR nuspec file ($nuspecfilname) does not exist
}

[xml]$nuspec = Get-Content $nuspecfilname

$url="https://www.nuget.org"
$version = $nuspec.package.metadata.version
$packagename = $nuspec.package.metadata.id 
$packagefilename = $packagename + "." + $version + ".nupkg"
$content_folder = join-path $packaging_folder "content"
$package_content = join-path $content_folder $packagename 
$lib_folder = join-path $packaging_folder "lib"
$tools_folder = join-path $packaging_folder "tools"


# Clean up any earlier packaging files
if (test-path $packaging_folder)
{
    remove-item -Recurse -Force $packaging_folder 
}

# Recreate the folder that will hold the contents
mkdir $package_content 

# If the lib folder does not exist, create it
# NOTE: nuget.exe requires the lib folder to exist 

if ( ! (test-path $lib_folder) ) 
{ 
    mkdir $lib_folder
}

# If the tools folder does not exist, create it
# NOTE: nuget.exe requires the tools folder to exist 

if ( ! (test-path $tools_folder) ) 
{ 
    mkdir $tools_folder
}

# Copy all the CS files except any that begin with "exclude"
robocopy $packagename  $package_content *.cs /xf exclude*.* 

# Create the NuGet Package
./nuget.exe pack $nuspecfilname -Verbose 


Write-Host "Publishing"
Write-host "param_pub": $param_pub
Write-host "pub_params": $pub_params
Write-host "publocal": $publocal
Write-host "pubremote": $pubremote

# Local publishing
if ($publocal)
{
    Write-Host "Publish locally"
    if (!(test-path $localfeed))
    {
        Write-Host ERROR $localfeed Does not exist
        Exit
    }

    Write-Host Copying to: $localfeed
    copy $packagefilename $localfeed

}

# Global publishing
if ($pubremote)
{
    Write-Host "Publish remote"
    Write-Host Publishing $packagefilename 
    .\nuget.exe push $packagefilename -Source $publicfeedurl
}

# Remove the packaging folder
remove-item -Recurse -Force $packaging_folder 
