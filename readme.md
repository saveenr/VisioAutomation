## Introduction
VisioAutomation makes it easier for developers to control Visio from .NET languages. In addition to simplifying common tasks, it makes it easier to build your own Visio add-ins and automation/scripting tools. Finally, it enables you to generate simple Visio XML (VDX) files without even having Visio installed.
Source

## Cloning
Git Clone url: https://github.com/saveenr/VisioAutomation/

## Tools

Download the tools here: https://visioautomation.codeplex.com/releases

## Visio Power Tools
Power Tools for Visio 2007 - stable but in maintenance mode
Power Tools for Visio 2010 - beta but in active development

## Visio PowerShell (VisioPS)
This is a PowerShell module for Automating Visio 2010.

Documentation here: http://sdrv.ms/11AWkp7


## Get the VisioAutomation via NuGet Packages

## Documentation

* Visio Automation 2010 - https://onedrive.live.com/?cid=1ff099edb1c7ebfa&id=1FF099EDB1C7EBFA%2137994&authkey=!AHhZWdAlU7LwJ6Y
* Visio Automation 2007 - https://onedrive.live.com/?cid=1ff099edb1c7ebfa&id=1FF099EDB1C7EBFA%211817
* Visio PowerShell - http://1drv.ms/18lQ7o5
* 


## Visio PowerShell (VisioPS) Release Notes

### Documentation is here http://sdrv.ms/11AWkp7

Screencast

http://vimeo.com/61329170
Files

For easy installation, download and run the MSI file.
If you want to manually install, a ZIP file is provided.
ChangeLog

Version 1.2.201

Regenerated the MSI Installer
Version 1.2.200

VisioPS should now work better when a Master is open for editing
New cmdlet: Open-VisioMaster
New cmdlet: Close-VisioMaster
New cmdlet: Format-VisioText

Version 1.1.25

fixed Get-VisioCustomProperties
Version 1.1.23

Fixed a logging bug - was dividing by zero when operations were done on zero shapes
Version 1.1.21

Cmdlets Get-VisioPageCell and Get-VisioShapeCell now support specific switch parameters for cells. This makes them work similarly to Set-VisioPageCell and Set-VisioShapeCell. You can mix using -Cells with any of the switch parameters for specific cells.
Cmdlets Get-VisioPageCell and Get-VisioShapeCell now support wildcards with tje -Cells parameter.

Get-VisioShapeCell -Cells *

Get-VisioShapeCell -Width -Cells *

Get-VisioShapeCell -Width -Cells Fill*

Get-VisioShapeCell -Width -Cells Fill,Lock
New-VisioDocument will create a new Visio application if needed (if there isn't an application bound to the session or if that application reference is no longer valid)

Version 1.1.20

Set-VisioPage replaced -Flags with -Direction for relative page navigation
New-VisioGridLayout now supports optional parameters to control the spacing between cells
Version 1.1.14

Fixed a bug that presented CustomProperties from being retrieved correctly
Version 1.1.13

For user convenience - cmdlets now return collections to the pipeline as a full collection instead of one object at a time
Version 1.1.12

Connect-VisioApplication no longer forgets which application it connected to
Version 1.1.12

Fixed a bug in Pie Slice rendering when slices were greater or equal to 180 degrees
Version 1.1.11

Now using new VisioAutomation library for increased performance when working with ShapeSheet
Version 1.1.10

This is the same as 1.1.9, except that the MSI installer is improved. It now supports upgrading from older versions.
Version 1.1.9

Fixed a bug that prevented VisioPS from drawing Org Charts correctly on Visio 2013
Last edited Oct 26, 2014 at 2:00 AM by saveenr, version 6

