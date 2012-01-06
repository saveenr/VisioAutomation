import clr
import System
import datetime
clr.AddReference("Microsoft.Office.Interop.Visio")
import Microsoft.Office.Interop.Visio

# sprint options
first_sprint_starts_on = datetime.date(2012,1,2)
num_sprints = 12
start_sprint_number = 7
days_in_sprint = 28

# rendering options
sprint_height = 1.0
sprint_width = 2.0
margin = 0.5
draw_vertical = False

# render based on options
IVisio = Microsoft.Office.Interop.Visio
visapp = IVisio.ApplicationClass()
doc = visapp.Documents.Add("")
page = visapp.ActivePage

if (draw_vertical) :
    page_width = margin*2 + sprint_width
    page_height = margin*2 + (num_sprints *sprint_height)
else :
    page_width = margin*2 + (num_sprints *sprint_width)
    page_height = margin*2 + sprint_height


page.PageSheet.CellsU["PageWidth"].FormulaU = page_width
page.PageSheet.CellsU["PageHeight"].FormulaU = page_height


for i in xrange(num_sprints) :
    sprintnumber = i+start_sprint_number
    sprintname = "Sprint %s" % ( sprintnumber ) 
    begin_date = first_sprint_starts_on + datetime.timedelta( i*days_in_sprint )
    end_date = begin_date + datetime.timedelta( days = (days_in_sprint-1) )
    print sprintname, begin_date, end_date

    if (draw_vertical) :
        x0 = margin
        x1 = margin + sprint_width
        y1 = (page_height - margin) - (i*sprint_height)
        y0 = y1 - sprint_height
    else :
        y0 = margin
        y1 = margin + sprint_height
        x1 = margin + (i*sprint_width)
        x0 = x1 + sprint_width
    
    shape = page.DrawRectangle(x0, y0, x1, y1)
    sprintlabel = sprintname + "\n" + begin_date.strftime("%m/%d") + " - " + end_date.strftime("%m/%d")
    shape.Text = sprintlabel
    

System.Console.ReadKey() 
