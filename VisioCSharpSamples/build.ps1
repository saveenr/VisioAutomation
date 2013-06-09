Set-StrictMode -Version 2 
$ErrorActionPreference = "Stop"

$mypath = $MyInvocation.MyCommand.path
$project_path = Resolve-Path ( Join-Path $MyInvocation.MyCommand.path ".." )
$samples_path = Join-Path $project_path "VisioCSharpSamples" 
$nuspecfilname = Join-Path $project_path "VisioCSharpSamples.nuspec"

Write-Host PROJ $project_path
Write-Host SAMPLES $samples_path
Write-Host NUSPEC $nuspecfilname


$packaging_folder = "d:\VisiosCSharpSamplesBuild\Packaging" 


Write-Host loading $nuspecfilname 

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

Write-Host Packaging Folder: $packaging_folder

if ( ! (test-path "d:\VisiosCSharpSamplesBuild") ) 
{
    mkdir "d:\VisiosCSharpSamplesBuild"
}

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
Write-Host from: $samples_path
robocopy $samples_path  $package_content *.cs /xf Program.cs 


# Create the NuGet Package

$nuget_exe = Join-Path $project_path "nuget.exe"
Write-Host NUGET $nuget_exe

if (!(test-path $nuget_exe))
{
    Write-Error "NuGET.EXE can't be found"
}


$dest_nuspec = Join-Path "d:\VisiosCSharpSamplesBuild" "VisioCSharpSamples.nuspec"
copy $nuspecfilname $dest_nuspec 

$old_location = Get-Location
Set-Location "d:\VisiosCSharpSamplesBuild"
&$nuget_exe pack $dest_nuspec -Verbose
Set-Location $old_location

# Remove the packaging folder
remove-item -Recurse -Force $packaging_folder 
