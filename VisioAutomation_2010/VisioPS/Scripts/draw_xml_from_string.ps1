Import-Module VisioPS

New-VisioApplication
New-VisioDocument

$xml = [xml]"<a><b><c/></b><d><e/><f><g/></f></d></a>"
$xml | Out-Visio

