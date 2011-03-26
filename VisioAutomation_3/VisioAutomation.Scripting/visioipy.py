import sys
import os
import clr
import System

# If you you are having trouble loading the modules, use the debug flag
# to help sort out where the problem might be
debug = False

visioipy_startup = os.environ.get( 'IRONPYTHONSTARTUP' , None )

# if the IRONPYTHONSTART env variable exists add it to sys.path
if ( visioipy_startup != None ) :
    visioipy_path = os.path.dirname( visioipy_startup )
    if (visioipy_path not in sys.path) :
        sys.path.append(visioipy_path)

if ( debug ) :
	print "IRONPYTHONSTARTUP =", visioipy_startup
	for i in xrange(len(sys.path)) :
		print "sys.path[%s] = \"%s\"" % (i, sys.path[i])
	print

# Get a reference to Visio
clr.AddReference( "Microsoft.Office.Interop.Visio" )
import Microsoft.Office.Interop.Visio

# Load the VisioAutomation assemblies
# Note: Adding this directory to syspath is needed for the imports of these DLLs to work
sys.path.append( os.path.join( sys.path[0] ,"") ) 
clr.AddReferenceToFileAndPath( System.IO.Path.Combine( visioipy_path, "VisioAutomation.dll" ) )
clr.AddReferenceToFileAndPath( System.IO.Path.Combine( visioipy_path, "VisioAutomation.Scripting.dll" ) )

import VisioAutomation
import VisioAutomation.Scripting

# Create aliases
IVisio = Microsoft.Office.Interop.Visio 
VA = VisioAutomation
VAS = VisioAutomation.Scripting

# Start a new Scripting session
if ( "visio" not in dir() ) : visio = VAS.Session()
assert( visio != None )

