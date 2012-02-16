#
# Sample code to draw a simple 1 month calendar in Visio using IronPython
# Demonstrates the use of the Python calendar module
#
# Author: Saveen Reddy
# Created: 2010-10-19
# Updated: 2010-10-10
#

import sys
import clr
import System
import calendar
clr.AddReference("Microsoft.Office.Interop.Visio")
import Microsoft.Office.Interop.Visio

# Some Preferences
#firstday = 0 # week starts on Monday
firstday = 6 # week starts on Sunday
target_year = 2010
target_month = 10
cellwidth = 1.0
cellheight = 0.5
textcolor_out_of_range_day = "rgb(180,180,180)"

# Build the calendar and get info about the target month
mycalendar = calendar.Calendar(firstday)
weeks = mycalendar.monthdatescalendar(target_year,target_month)
first_week = weeks[0]

# Prepare Visio
IVisio = Microsoft.Office.Interop.Visio
visapp = IVisio.ApplicationClass()
doc = visapp.Documents.Add("")
page = doc.Pages.Add()

# Draw the days
for wi,w in enumerate(weeks):
    for di,d in enumerate(w):
        row = len(weeks)-1-wi
        x0, y0 = di*cellwidth, row*cellheight
        shape = page.DrawRectangle(x0,y0,x0+cellwidth,y0+cellheight)
        shape.Text = d.day
        if (d.month!=target_month) :
            cell = shape.Cells[ "Char.Color" ]
            cell.FormulaU = textcolor_out_of_range_day

# draw the header with names of the days
for di in xrange(7):
    x0, y0 = di*cellwidth, len(weeks)*cellheight
    x1, y1  = x0 + cellwidth, y0 + cellheight
    shape = page.DrawRectangle(x0,y0,x1,y1)
    shape.Text = calendar.day_name[ first_week[di].weekday() ]