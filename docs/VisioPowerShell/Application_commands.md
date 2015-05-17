# Creating a new Application

Note that New-VisioApplication does NOT return an application object.

	$visapp = New-VisioApplication
	# $visapp will always be null!

# Getting the bound application instance
You can retrieve the bound Visio application instance by using Get-VisioApplication

	$visapp = Get-VisioApplication

Checking if the Bound Application is still valid
Sometimes you'll need to perform an action only if a valid instance of Visio is connected to VisioPS. The easy way to check is to use the Test- VisioApplication cmdlet.

	if (Test-VisioApplication)
	{
	    # the bound application is valid
	}
	else
	{
	    # the bound application is not valid
	}


