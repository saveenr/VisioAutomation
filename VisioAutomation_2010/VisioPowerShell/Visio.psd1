# Manifest for "Visio" PowerShell module (VisioPS)
#

@{

# Script module or binary module file associated with this manifest.
# RootModule = 'VisioPS.dll' - Commented this out because having RootModule defined causes the module to fail to load with PowerShell 2.0
ModuleToProcess = 'VisioPS.dll' # Use ModuleToProcess instead of RootModule because it works for both PowerShell 2.0 and 3.0

# Version number of this module.
ModuleVersion = '3.0.1'

# ID used to uniquely identify this module
GUID = 'd2d6f65b-2eee-4397-98ee-94ff7930051c'

# Author of this module
Author = 'Saveen Reddy'

# Company or vendor of this module
CompanyName = ''

# Copyright statement for this module
Copyright = 'Saveen Reddy'

# Description of the functionality provided by this module
Description = 'Visio PowerShell - Automatation cmdlets for Visio version 2010 and above'

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
# RequiredAssemblies = @()

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

# Cmdlets to export from this module
CmdletsToExport = '*'

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
        #ReleaseNotes = '* Bug Fix for TokenCache initialization when importing a context'

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
