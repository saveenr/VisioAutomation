# PURPOSE
# -------
# Loads the module directly from bin/debug
#
# This the fastest way to test the module without installing the module
#

$visio_psd1 = join-path $PSScriptRoot ".\bin\Debug\Visio.psd1"
Import-Module $visio_psd1 

# Get a new document ready
New-VisioApplication
New-VisioDocument


$page_cells = New-VisioPageCells

$page_cells.PageHeight = "5 in"
$page_cells.PageHeight = "10 in"
$page_cells.PrintLeftMargin = "0 in"

New-VisioPage -Name "HelloWorld" -Cells $page_cells -Verbose



