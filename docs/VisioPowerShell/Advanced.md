
# Using the VisioAutomation Assembly
Sometimes you may need to directly use types in the VisioAutomation assembly. Use the Get-VisioAutomationAssembly to get the assembly object. Then load the assembly from disk. Then you can use the New-Object cmdlet to create VisioAutomation objects directly.
Below is an example of creating a new VisioAutomation.Drawing.Point object and new 
	VisioAutomation.Drawing.Rectangle.
	$asm = Get-VisioAutomationAssembly
	
	[Reflection.Assembly]::LoadFile( $asm.Location )
	
	$p = New-Object VisioAutomation.Drawing.Point(1,2)
	$r = New-Object VisioAutomation.Drawing.Rectangle(1,2,3,4)
	$pinx_src = [VisioAutomation.ShapeSheet.SRCConstants]::PinX

# Getting the current ScriptingSession

Most of the logic in VisioPS is actually implemented by a layer called VisioAutomation.Scripting. Specifically there is a ScriptingSession object that takes care of most of the hard work of implementing an interactive tool that interacts with Visio. 

VisioPS is little more than a PowerShell-oriented wrapper around the ScriptingSession object.

Sometimes you will find it useful to directly interact with the ScriptingSession. You can get the current SciptingSession object by using the Get-VisioAutomationScriptingSession cmdlet.

	$ss = Get-VisioAutomationScriptingSession

# Creating a new Visio Application COM Object

	$application = New-Object -ComObject Visio.Application

