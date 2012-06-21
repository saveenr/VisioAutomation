./pubinfo.ps1
remove-item -Recurse -Force .\Packaging\content\VisioCSharpSamples
mkdir .\Packaging\content\VisioCSharpSamples
robocopy .\VisioCSharpSamples .\Packaging\content\VisioCSharpSamples *.cs /xf exclude*.* 
./nuget.exe pack Package.nuspec -Verbose -Version $global:version

