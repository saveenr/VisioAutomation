


Installing the Module for Testing on your Machine
-----
Run Install.ps1. It will mannyally install the module for the current user.


Installing the Module as a user would.
-----
This means publishing the module to PowerShell gallery and then installing
the module via the Install-Module cmdlet.

To install for the current user

    Install-Package -Name -Scope CurrentUser

To install for all users, run the following as Adminsitrator

    Install-Package -Name 


Publishing the package
-----

1. First Compiler the project
2. Install the module via install.ps1
3. Run Publish-Module -Name Visio -NuGetApiKey ***************

The NuGetApiKey can be retrieved by signing into PowerShell gallery org
 

