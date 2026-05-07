# Manifest for "Visio" PowerShell module (VisioPS)
#

@{

# Script module or binary module file associated with this manifest.
# RootModule = 'VisioPS.dll' - Commented this out because having RootModule defined causes the module to fail to load with PowerShell 2.0
ModuleToProcess = 'VisioPS.dll' # Use ModuleToProcess instead of RootModule because it works for both PowerShell 2.0 and 3.0

# Version number of this module.
ModuleVersion = '4.7.2'

# ID used to uniquely identify this module
GUID = 'd2d6f65b-2eee-4397-98ee-94ff7930051c'

# Author of this module
Author = 'SevenPens'

# Company or vendor of this module
CompanyName = ''

# Copyright statement for this module
Copyright = 'SevenPens'

# Description of the functionality provided by this module
Description = 'Visio PowerShell - Automation cmdlets for Visio version 2010 and above'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '2.0'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of the .NET Framework required by this module
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module
CLRVersion = '4.0'

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
RequiredAssemblies = @(
"VisioAutomation.dll",
"VisioAutomation.Models.dll",
"VisioPS.dll",
"VisioScripting.dll",
"Microsoft.Msagl.dll",
"GenTreeOps.dll"
)


# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
TypesToProcess = @('Visio.Types.ps1xml')

# Format files (.ps1xml) to be loaded when importing this module
FormatsToProcess = @(  )

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
NestedModules = @()

# Functions to export from this module
FunctionsToExport = '*'

# Cmdlets to export from this module.
# Enumerated explicitly (rather than '*') to satisfy the PSGallery publish-time
# best-practice check and to skip module-load reflection over the binary. When
# adding a new cmdlet, append it here. Drift between this list and the cmdlets
# VisioPS.dll actually exports is caught by VisioPS_Manifest_Tests in
# VTest.PowerShell (every test run) and as a defense-in-depth check by
# publish-psmodule.yml (every publish run).
CmdletsToExport = @(
    'Close-VisioApplication',
    'Close-VisioDocument',
    'Connect-VisioShape',
    'Copy-VisioPage',
    'Copy-VisioShape',
    'Export-VisioPage',
    'Export-VisioShape',
    'Format-VisioPage',
    'Format-VisioShape',
    'Format-VisioWindow',
    'Get-VisioApplication',
    'Get-VisioClient',
    'Get-VisioControl',
    'Get-VisioCustomProperty',
    'Get-VisioDocument',
    'Get-VisioHyperlink',
    'Get-VisioLockCells',
    'Get-VisioMaster',
    'Get-VisioPage',
    'Get-VisioPageCells',
    'Get-VisioShape',
    'Get-VisioShapeCells',
    'Get-VisioText',
    'Get-VisioUserDefinedCell',
    'Import-VisioModel',
    'Join-VisioShape',
    'Lock-VisioShape',
    'Measure-VisioPage',
    'Measure-VisioShape',
    'New-VisioApplication',
    'New-VisioContainer',
    'New-VisioControl',
    'New-VisioDocument',
    'New-VisioHyperlink',
    'New-VisioPage',
    'New-VisioPageCells',
    'New-VisioPoint',
    'New-VisioRectangle',
    'New-VisioShape',
    'New-VisioShapeCells',
    'Open-VisioDocument',
    'Out-VisioApplication',
    'Redo-VisioApplication',
    'Remove-VisioControl',
    'Remove-VisioCustomProperty',
    'Remove-VisioHyperlink',
    'Remove-VisioPage',
    'Remove-VisioShape',
    'Remove-VisioUserDefinedCell',
    'Save-VisioDocument',
    'Select-VisioDocument',
    'Select-VisioPage',
    'Select-VisioShape',
    'Set-VisioCustomProperty',
    'Set-VisioPageCells',
    'Set-VisioShapeCells',
    'Set-VisioText',
    'Set-VisioUserDefinedCell',
    'Split-VisioShape',
    'Test-VisioApplication',
    'Test-VisioDocument',
    'Test-VisioShape',
    'Undo-VisioApplication',
    'Unlock-VisioShape'
)

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = '*'

# List of all modules packaged with this module.
ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = 'Visio'

        # A URL to the license for this module.
        LicenseUri = 'https://github.com/saveenr/VisioAutomation/blob/master/LICENSE.txt'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/saveenr/VisioAutomation'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        ReleaseNotes = 'See https://github.com/saveenr/VisioAutomation/blob/master/VisioAutomation_2010/VisioPowerShell/CHANGELOG.md'

        # Prerelease string of this module
        # Prerelease = ''

        # Flag to indicate whether the module requires explicit user acceptance for install/update
        # RequireLicenseAcceptance = $false

        # External dependent modules of this module
        # ExternalModuleDependencies = @()

    } # End of PSData hashtable

 } # End of PrivateData hashtable


# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}
