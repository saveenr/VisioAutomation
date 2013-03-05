
Function CreateSubFolder( $path, $name) 
{
    $new_folder = Join-Path $path $name
    if (!(Test-Path $new_folder ))
    {
        New-Item -Path $new_folder -ItemType Directory | Out-Null
    }
    Write-Output $new_folder
} 

Function Install-ModuleForUser
{ 
    Param( 
        [Parameter(Mandatory=$true)]
        [String] $Folder,

        [Parameter(Mandatory=$true)]
        [String] $Name

    ) 

    Process 
    {
        $destfolder = Join-Path $env:USERPROFILE "Documents"
        $wps_folder = CreateSubFolder $destfolder "WindowsPowerShell"
        $modules_folder = CreateSubFolder $wps_folder "Modules"
        $mod_folder = CreateSubFolder $modules_folder $Name

        Robocopy `"$Folder`" `"$mod_folder`" /MIR                    
    }
}

$localdll = Join-Path $MyInvocation.MyCommand.Path "..\..\bin\debug"
$localdll = Resolve-Path $localdll

Install-ModuleForUser $localdll "VisioPS"