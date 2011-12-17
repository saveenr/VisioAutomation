#draws a inch grid on a 8.5x11 drawing
import-module Visiops.dll
New-VisioApplication
New-drawing
set-pagelayout -width 8.5 -height 11


# Draw the vertical lines
for ($x=0; $x -le 8.5; $x++) { Draw-Line $x 0 $x 11 }

# Draw the horzontal lines
for ($y=0; $y -le 11; $y++) { Draw-Line 0 $y 8.5 $y }
