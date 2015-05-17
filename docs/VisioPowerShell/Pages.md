# Pages

## Note
Visio Documents can't have zero pages. If you create a new document you are guaranteed to have a page.

## Create a new page
	New-VisioPage

## Creating 20 new pages

	foreach ($i in (0..19)) { New-VisioPage }

## Create a new page with a specific size and anme
	New-VisioPage -Width 10.0 -Height 4.0 -Name "MyPage1"

## Get All Pages from Active Document
	
	$pages = Get-VisioPage

## Get All Pages from Active Document that have a specific name

	$pages = Get-VisioPage â€œPage-1â€

Wildcards are supported

	$pages = Get-VisioPage *2

# Get the Active Page from the Active Document

	$page = Get-VisioPage -ActivePage

# Setting the Page Size

There are two ways to do this.

Option 1: Setting the ShapeSheet cells directly

	Set-VisioPageCell -PageWidth 10.0 -PageHeight 4.0

Option 2: Use Resize-VisioPage

	Resize-VisioPage -Width 10.0 -Height 4.0

# Resize the Page to fit its contents
	
	Resize-VisioPage -FitContents   

You may want some extra padding
	Resize-VisioPage -FitContents -BorderWidth 1.0 -BorderHeight 1.0

# Page Orientation and Background

	Set-VisioPageLayout -Orientation Portrait
	
	Set-VisioPageLayout -Orientation Landscape


# Set the Active Page
	Set-VisioPage -Name "Mypage"

	Set-VisioPage -Page $p

Relative navigation is also possible

	Set-VisioPage -Direction First
	Set-VisioPage -Direction Last
	Set-VisioPage -Direction Next
	Set-VisioPage -Direction Previous

# Duplicate a Page

	Invoke-VisioDuplicate-Page

To give the new page a specific name use -Name
To place the new page in a different document use -ToDocument

# Deletes the active page

	Remove-VisioPage

# Delete specific pages

	$pages = Get-VisioPage

	Remove-VisioPage $pages[0] # deletes the first page

	$pages = Get-VisioPage

	Remove-VisioPage $pages[0],$pages[3] # deletes the first and fourth page

# Delete all pages
	$pages = Get-VisioPage

	Remove-VisioPage $pages

Delete all pages that have "2" in their name
	Get-VisioPage *2* | Remove-VisioPage

