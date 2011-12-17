#
# Draw some rectangles and drop some shapes
#

import-module .\VisioPS.Dll

$visapp = New-VisioApplication
$d1 = New-Drawing

$basic_stencil = Open-Stencil basic_u.vss
$master = Get-Master "Rectangle" "Basic_U.VSS"

$s1 = Draw-Rectangle 0 0 1 1
$s2 = Draw-Rectangle 3 2 4 5
$s3 = Drop-Master $master 3,3
 