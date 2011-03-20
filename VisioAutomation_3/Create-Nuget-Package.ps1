
# start package metadata
$title = "VisioAutomation"
$version = "3.2.x"
$author = "Saveen Reddy"
$description = "The VisioAutomation simplified the control of the Visio Application via .NET Languages"
$projecturl = "http://visioautomation.codeplex.com"
# end package metadata





function make_directory( [string] $p )
{
    if (-not (test-path $p)) 
    {
        New-Item $p -type directory
    }
}

function make_package( [string] $libsrc, [string] $extra )
{
    $mydocs = [Environment]::GetFolderPath("MyDocuments")
    $packagepath = join-path $mydocs ( $title + $version )
    write-host $extra    

    write-host $packagepath
    make_directory $packagepath

    $fullpath = join-path $packagepath ($title + $extra + "-" + $version)
    write-host $fullpath
    
    make_directory $fullpath
 
    $libpath = join-path $fullpath lib
    
    write-host $libsrc
    xcopy $libsrc $libpath /I /Y

    $fname = join-path $fullpath "package.nuspec"
    
    write-host $fullpath
    write-host $libpath
    write-host $fname
    $template = @"
<package>
  <metadata>
    <id>$title$extra</id>    
    <title>$title$extra</title>
    <version>$version</version>
    <authors>$author</authors>
    <description>$description</description>
    <projectUrl>$projecturl</projectUrl>
  </metadata>
  <files>
      <file src="lib\*.dll" target="lib" /> 
  </files>
</package>
"@

    $template | Out-File $fname -encoding UTF8
}

make_package "VisioAutomation.Scripting\bin\Debug" ".Debug"
make_package "VisioAutomation.Scripting\bin\Release" ""
