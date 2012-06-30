
$localfeed ="D:\saveenr\live-mesh\code\mynugetfeed" 
$publicfeedurl="https://www.nuget.org"

$publocal = $true
$pubremote = $false

$url="https://www.nuget.org"


add-type @"
public class PackageRecord
{
   public string Name;
   public string Version;
}
"@

$pkg_va2010 = new-object PackageRecord
$pkg_va2010.Name = "VisioAutomation2010"
$pkg_va2010.Version = "1.0.0.0"

$pkg_vavdx = new-object PackageRecord
$pkg_vavdx.Name = "VisioAutomation.VDX"
$pkg_vavdx.Version = "1.0.0.0"

$packagerecords = @($pkg_va2010, $pkg_vavdx)

foreach ($pk in $packagerecords)
{
    
    $packagefilename = $pk.Name + "." + $pk.version + ".nupkg"
    $nuspec = $pk.Name + ".nuspec"

    Write-Host ">>>>>>>>>", $pk.Name
    Write-Host ">>>>>>>>>", $nuspec 

    # Create the NuGet Package
    ./nuget.exe pack $nuspec -Verbose -Version $version
    
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

