# Creating a New Document

	$d = New-VisioDocument

The cmdlet will create a new Visio Application object if one is not needed

# Get All Documents
	 
	$alldocs = Get-VisioDocument

# Get All Documents based on name
To find a specific document Get-VisioDocument "DocumentFoo" To find all documents 

	$alldocs = Get-VisioDocument *

## Wildcards are supported.
# Returns all documents with a "3" in their name

	$alldocs = Get-VisioDocument *3*

# Get the Active Document

	Get-VisioDocument -ActiveDocument

# Closing a Document

	Close-VisioDocument

If the document has been modified it wonâ€™t be closed. Instead youâ€™ll be prompted to close it or not. To force it to close use -Force.

	Close-VisioDocument -Force

You can specify the document object to close 
	Close-VisioDocument -Documents $doc1
You can specify multiple to close 
	Close-VisioDocument -Documents doc1,doc2 -Force

# Force all documents to close

	$docs = Get-VisioDocument
	Close-VisioDocuments $docs -Force

# Setting the Active Document
	
	Set-VisioDocument $doc
Or you can name the document
	
	Set-VisioDocument "Drawing5"

# Checking is the Active Document is valid
Sometimes you'll need to perform an action only if a drawing is open. The easy way to check is to use Test-VisioDocument

	If (Test-VisioDocument)
	{
	    # do something
	}
	else
	{
	    # do something else
	}

# Save

	Save-VisioDocument

# Save As
	
	Save-VisioDocument "d:\foo.vsd"

Load a Document

	Open-VisioDocument "d:\foo.vsd" 


