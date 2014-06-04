Import-Module VisioPS

New-VisioApplication
New-VisioDocument


$xml = new-object System.Xml.XmlDocument
$xml.load("./test.csproj")
Out-Visio -XmlDocument $xml -Verbose
