# Visio PowerShell User Guide
Author: Saveen Reddy
Author email: saveenr@live.com
Updated: 2015/05/17
Version this Document covers: VisioPS 1.2.205

# Introduction
I wrote this PowerShell module (abbreviated as VisioPS) to simplify automating common tasks with Visio 2010 and above.  


# Demonstration
This presentation at the PowerShell Summit 2014 demonstrates the basics: 
https://vimeo.com/94408016

# Performance
Whenever possible this module communicates and controls Visio using high-performance techniques.

# PowerShell-Friendly
This isn't just some object-model wrapper on top of Visio. It was designed from the beginning to work in a way that makes it convenient for PowerShell fans.

# Visio Version Support
* This module works with Visio2010 and Visio2013.
* Because VisioPS is developed using Visio2010 it doesn't have support for any features unique to Visio2013.

# PowerShell Version Support
* Both PowerShell 3.0 and PowerShell 2.0 are supported
* It's easiest if you just use PowerShell 3.0; If you use PowerShell 2.0, you must configure PowerShell 2.0 to use.NET Framework 4.0

# Installing PowerShell 3.0
* If you have Windows 7 or below: Install Windows Management Framework 3.0 (http://www.microsoft.com/en-us/download/details.aspx?id=34595). WMF 3.0 includes PowerShell 3.0
* If you have Windows 8: You don't have to do anything. PowerShell 3.0 comes with Windows 8.

# Using PowerShell 2.0
* VisioPS is a managed assembly built using .NET Framework 4.0, however PowerShell 2.0 by default runs using .NET Framework 2.0
* In order to use PowerShell 2.0, you'll need to configure PowerShell 2.0 to use .NET Framework 4.0. The instructions are here: http://viziblr.com/news/2012/5/16/the-easy-way-to-run-powershell-20-using-net-framework-40.html

# Installation
* Go here to find the MSI installer : http://visioautomation.codeplex.com/releases
* Click on the MSI to start the installation.
* VisioPS will be placed in a <MyDocuments>/Windows PowerShell/Modules/Visio
Getting Started


# Importing the Module

	Import-Module Visio


# Finding All the Visio Commands
There are a lot of commands :-) To see a list of them use the Get-Command cmdlet as shown below

	Get-Command -Module Visio | Select Name
Output:

	Name                                                                                               
	----                                                                                               
	Close-VisioApplication                                                                             
	Close-VisioDocument                                                                                
	Close-VisioMaster                                                                                  
	Copy-VisioPage                                                                                     
	Copy-VisioShape                                                                                    
	Export-VisioPage                                                                                   
	Export-VisioSelectionAsXHTML                                                                       
	Format-VisioShape                                                                                  
	Format-VisioText                                                                                   
	Get-VisioApplication                                                                               
	Get-VisioClient                                                                                    
	Get-VisioConnectionPoint                                                                           
	Get-VisioControl                                                                                   
	Get-VisioCustomProperty                                                                            
	Get-VisioDirectedEdge                                                                              
	Get-VisioDocument                                                                                  
	Get-VisioLayer                                                                                     
	Get-VisioMaster                                                                                    
	Get-VisioPage                                                                                      
	Get-VisioPageCell                                                                                  
	Get-VisioShape                                                                                     
	Get-VisioShapeCell                                                                                 
	Get-VisioText                                                                                      
	Get-VisioUserDefinedCell                                                                           
	Import-VisioModel                                                                                  
	New-VisioApplication                                                                               
	New-VisioAreaChart                                                                                 
	New-VisioBarChart                                                                                  
	New-VisioBezier                                                                                    
	New-VisioConnection                                                                                
	New-VisioControl                                                                                   
	New-VisioDirectedGraph                                                                             
	New-VisioDocument                                                                                  
	New-VisioGridLayout                                                                                
	New-VisioGroup                                                                                     
	New-VisioLine                                                                                      
	New-VisioMaster                                                                                    
	New-VisioNURBS                                                                                     
	New-VisioOrgChart                                                                                  
	New-VisioOval                                                                                      
	New-VisioPage                                                                                      
	New-VisioPieChart                                                                                  
	New-VisioPolyLine                                                                                  
	New-VisioRectangle                                                                                 
	New-VisioShape                                                                                     
	Open-VisioDocument                                                                                 
	Open-VisioMaster                                                                                   
	Out-Visio                                                                                          
	Redo-Visio                                                                                         
	Remove-VisioControl                                                                                
	Remove-VisioCustomProperty                                                                         
	Remove-VisioGroup                                                                                  
	Remove-VisioPage                                                                                   
	Remove-VisioShape                                                                                  
	Remove-VisioUserDefinedCell                                                                        
	Resize-VisioPage                                                                                   
	Save-VisioDocument                                                                                 
	Select-VisioShape                                                                                  
	Set-VisioCustomProperty                                                                            
	Set-VisioDocument                                                                                  
	Set-VisioPage                                                                                      
	Set-VisioPageCell                                                                                  
	Set-VisioPageLayout                                                                                
	Set-VisioShapeCell                                                                                 
	Set-VisioShapeSheet                                                                                
	Set-VisioText                                                                                      
	Set-VisioUserDefinedCell                                                                           
	Set-VisioWindowSize                                                                                
	Set-VisioZoom                                                                                      
	Test-VisioApplication                                                                              
	Test-VisioDocument                                                                                 
	Test-VisioSelectedShapes                                                                           
	Undo-Visio              

This output is ordered by Verb, it might be useful for you to organize by Noun as shown below

	Get-Command -Module Visio | Sort-Object Noun | Select Name

Getting Help for a Cmdlet

Get-Help allows you to find basic information about the syntax for each cmdlet

	Get-Help Set-VisioText

Output:

	NAME
	    Set-VisioText
	    
	SYNTAX
	    Set-VisioText [-Text] <string[]> [-Shapes <Shape[]>]  [<CommonParameters>]
	    
	
	ALIASES
	    None
	    
	
	REMARKS
	    None


# Troubleshooting with the -Verbose flag
If you are having trouble with the cmdlets, please do use the –Verbose flag. Most of the cmdlets will show additional information that will help you understand what is happening.

# Hello World
Let's cover a simple example that illustrates a lot about how to use VisioPS. First Start PowerShell

## Getting a new Application
Now create a new Visio Application

	New-VisioApplication

NOTE: For a technical reason New-VisioApplication does not return the application object. You can retrieve it later though with Get-VisioApplication

## Creating a New Document
By default an application has no documents, so we create one

	New-VisioDocument

NOTES:

* We didn't really need to call New-VisioApplication before New-VisioDocument because New-VisioDocument will create an application if needed
* A Document in Visio always has at least one page

## Dropping Shapes
Typically Visio documents are drawn by dropping Masters from Stencils. So let's start by loading the Basic Shapes stencil

	$basic_u = Open-VisioDocument basic_u.vss

And now find the master we are interested in which is called "Rectangle"
	$master = Get-VisioMaster "Rectangle" $basic_u

And now we create a shape using the Rectangle master

	$shape = New-VisioShape $master 3,3

## Modifying a Shape

Now let's set the text of the shape
	Set-VisioText "Hello World"

## The Full Hello World Script

The full script looks like this

	Import-Module Visio
	
	New-VisioDocument
	
	$basic_u = Open-VisioDocument basic_u.vss
	$master = Get-VisioMaster "Rectangle" $basic_u
	$shape = New-VisioShape $master 3,3
	
	Set-VisioText "Hello World"



## New-VisioApplication
Note that New-VisioApplication does NOT return an application object. 

	$visapp = New-VisioApplication
	# $visapp will always be null!

## Getting the bound application instance
You can retrieve the bound Visio application instance by using Get-VisioApplication

	New-VisioApplication
	$visapp = Get-VisioApplication

# Context Sensitivity

## Context Sensitive cmdlets
Notice that in these commands, we never had to identify the target of the operation.

For example, we didn't have to saw which document to load the stencil into - VisioPS will simply use the currently active document

Likewise we didn't have to identify which page to put the shape on, VisioPS will assume the currently active page

When we called Set-VisioText text at the end we didn't specify which shapes to target, VisioPS will automatically target the currently selected shapes

## Overriding Context Sensitivity

Context sensitive cmdlets make it simpler to automate Visio interactively, however sometimes you don't care about that and what to specify the targets explicitly. For that reason many cmdlets have parameters that let you specify exactly which Document, Page, or Shape you want to target. These will be visible as the following parameters on a number of cmdlets.

* -Documents
*	-Pages
*	-Shapes

In the example above we wrote:

	Set-VisioText "Hello World"

We could have also written:

	Set-VisioText "Hello World" -Shapes $shape
Which would have accomplished the same thing.


