Set-StrictMode -Version 2 
$ErrorActionPreference = "Stop"

function mkdir_safe($dir)
{
    if ( ! (test-path $dir) ) 
    { 
        mkdir $dir
    }
}

function clean_dir_safe($dir)
{
    if (test-path $packaging_folder)
    {
        remove-item -Recurse -Force $packaging_folder 
    }
}


$mypath = $MyInvocation.MyCommand.path
$project_path = Resolve-Path ( Join-Path $MyInvocation.MyCommand.path ".." )
$samples_path = Join-Path $project_path "VisioCSharpSamples" 
$nuspecfilname = Join-Path $project_path "VisioCSharpSamples.nuspec"

$nuget_exe = Join-Path $project_path "nuget.exe"
Write-Host NUGET $nuget_exe

if (!(test-path $nuget_exe))
{
    Write-Error "NuGET.EXE can't be found"
}

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

mkdir_safe "d:\VisiosCSharpSamplesBuild"

# Clean up any earlier packaging files
clean_dir_safe $packaging_folder

# Recreate the folder that will hold the contents
mkdir_safe $package_content 
mkdir_safe $lib_folder # NOTE: nuget.exe requires the lib folder to exist 
mkdir_safe $tools_folder # NOTE: nuget.exe requires the tools folder to exist 


# Copy all the CS files except any that begin with "exclude"
Write-Host from: $samples_path
robocopy $samples_path  $package_content *.cs /xf Program.cs 

# ------------------------
# Copy the NUSPEC file over
# ------------------------

$dest_nuspec = Join-Path "d:\VisiosCSharpSamplesBuild" "VisioCSharpSamples.nuspec"
copy $nuspecfilname $dest_nuspec 

# ------------------------
# Create the NuGet Package
# ------------------------

Push-Location d:\VisiosCSharpSamplesBuild
try
{
    &$nuget_exe pack $dest_nuspec -Verbose 
}
finally
{
    Pop-Location
}

# Remove the packaging folder
clean_dir_safe $packaging_folder 
