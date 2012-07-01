
$localfeed ="D:\saveenr\live-mesh\code\mynugetfeed" 
$publicfeedurl="https://www.nuget.org"

$publocal = $true
$pubremote = $true

$url="https://www.nuget.org"


add-type @"
public class PackageRecord
{
   public string Name;
}
"@

$pkg_va2010 = new-object PackageRecord
$pkg_va2010.Name = "VisioAutomation2010"

$pkg_va2007 = new-object PackageRecord
$pkg_va2007.Name = "VisioAutomation2007"

$pkg_vavdx = new-object PackageRecord
$pkg_vavdx.Name = "VisioAutomation.VDX"

$packagerecords = @($pkg_va2010, $pkg_va2007, $pkg_vavdx)

foreach ($pk in $packagerecords)
{
    $nuspec = $pk.Name + ".nuspec"
    $nuspecdata = [xml](get-content $nuspec)
    $version = $nuspecdata.package.metadata.version

    Write-Host ">>>>>>>>>", $pk.Name
    Write-Host ">>>>>>>>>", $nuspec 
    Write-Host ">>>>>>>>>", $version

    $packagefilename = $pk.Name + "." + $version + ".nupkg"


    # Create the NuGet Package
    ./nuget.exe pack $nuspec -Verbose 

    if ( !($lastexitcode -eq 0) )
    {
        Write-Host "Error running nuget"
        exit 1
    } 
        
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
  
 
}

