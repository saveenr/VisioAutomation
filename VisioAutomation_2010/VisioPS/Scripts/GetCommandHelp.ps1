Import-Module ..\bin\debug\VisioPS.dll

$cmds = Get-Command -Module VisioPS

$n = 0
foreach ($cmd in $cmds) 
{ 
    Write-Host; "--------------------"; 
	get-help $cmd.Name 
    $n++
} 

Write-Host Found $n commands in Module
