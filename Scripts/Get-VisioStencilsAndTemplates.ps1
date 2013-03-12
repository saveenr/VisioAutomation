Function Get-VisioStencilsAndTemplates
{ 
    Param( 
        [Parameter(Mandatory=$false)]
        $Versions = @("2007", "2010"),

        [Parameter(Mandatory=$false)]
        $Languages = @("en")

    ) 

    Process 
    {
        $lang_to_id = @{
                      "en"="1033"
                      }

        $ver_to_path = @{
                       "2007" = "C:\Program Files (x86)\Microsoft Office\Office12" ;
                       "2010" = "C:\Program Files (x86)\Microsoft Office\Office14\Visio Content"
                       }

        $print_masters = $true

        foreach ($ver in $Versions)
        {
            foreach ($lang in $Languages)
            {
                Write-Host "LANAG" $lang
                $stencil_path = join-path $ver_to_path[$ver] $lang_to_id["en"] 

                if (!(test-path $stencil_path ))
                {
                    Continue;
                }
                Write-Host "Stencilpath " $stencil_path

                $vss_files = Get-ChildItem $stencil_path "*.vss" -File
                $vst_files = Get-ChildItem $stencil_path "*.vst" -File

                Write-Output $vss_files
                Write-Output $vst_files
            }
        }

    }
}


Get-VisioStencilsAndTemplates

