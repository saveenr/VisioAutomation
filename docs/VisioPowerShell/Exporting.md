# Export Page as an Image

	Export-VisioPage "d:\foo.png"

# Export Each Page to an separate Image

	Export-VisioPage "d:\foo.png" -AllPages

Will create a PNG for each page. The name will be of the form foo_Page_N.png

# Export Selection as XHTML (using SVG)

	Export-VisioSelectionAsXHTML "d:\foo.xhtml" E:_Docs_Documents.md
