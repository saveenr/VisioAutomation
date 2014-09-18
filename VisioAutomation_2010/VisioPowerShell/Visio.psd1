# Module manifest for module 'Visio PowerShell Module (VisioPS)'
#
# HISTORY
# -------
# 2014/05/23 Added additional metadata
# 2014/05/14 Renamed module from "VisioPS" to "Visio"
# 2013/08/06 Moved VisioPS.dll moved to RootModule
# 2012/02/16 Updated PowerShellVersion and Copyright
# 2012/08/08 Initial version
#

@{

# Script module or binary module file associated with this manifest.
# RootModule = 'VisioPS.dll' - Commented this out because having RootModule defined causes the module to fail to load with PowerShell 2.0
ModuleToProcess = 'VisioPS.dll' # Use ModuleToProcess instead of RootModule because it works for both PowerShell 2.0 and 3.0

# Version number of this module.
ModuleVersion = '1.2.200'

# ID used to uniquely identify this module
GUID = 'd2d6f65b-2eee-4397-98ee-94ff7930051c'

# Author of this module
Author = 'Saveen Reddy'

# Company or vendor of this module
CompanyName = ''

# Copyright statement for this module
Copyright = '(c) 2014 Saveen Reddy'

# Description of the functionality provided by this module
Description = 'Automate Microsoft Visio 2010 or Visio 2013'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '2.0'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of the .NET Framework required by this module
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module
# CLRVersion = ''

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
# PrivateData = ''

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}

