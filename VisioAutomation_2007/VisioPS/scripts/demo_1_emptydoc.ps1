#
# Starts a new application, creates a new drawing, and loads the Basic Shapes stencil
#

import-module .\VisioPS.Dll

$visapp = New-VisioApplication
$d1 = New-Drawing

$basic_stencil = Open-Stencil basic_u.vss
