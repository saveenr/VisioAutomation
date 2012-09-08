$script_filename = $MyInvocation.MyCommand.path 
$script_path = Split-Path $script_filename

$bin_path = join-path $script_path "binaries"

$bin_path

$ilmerge_exe = "C:\Program Files (x86)\Microsoft\ILMerge\ILMerge.exe"
$out_file = join-path $script_path "XamlTuneMerged.dll"
$primary_assembly = join-path $bin_path "XamlTuneConverter.dll"

& $ilmerge_exe /out:$out_file $primary_assembly /lib:$bin_path /v4 /target:library /log:merge.log /internalize