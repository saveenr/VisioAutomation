import clr 
import System 

clr.AddReference("System.Drawing") 
import System.Drawing 


# Get the colors from the bitmap 
bmp = System.Drawing.Bitmap( "D:\\swatch.png" ) 
pixels = list(set([ bmp.GetPixel(x,y) for x in xrange( bmp.Width ) for y in xrange( bmp.Height) ])) 

# Start Visio 
clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 
IVisio = Microsoft.Office.Interop.Visio 
visapp = IVisio.ApplicationClass() 
doc = visapp.Documents.Add("") 
page = visapp.ActivePage 


# Draw the shapes 
coords = [ (i%23,i/23) for i in xrange( len(pixels) ) ] 
length = 0.25 
rects = [ ( x*length, y*length, (x+1)*length, (y+1)*length ) for x,y in coords ] 
shapes = [ page.DrawRectangle( *rect ) for rect in rects ] 


#Adjust the page 
page.ResizeToFitContents() 


# Set the colors 
formulas = [ "=RGB({0},{1},{2})".format(p.R, p.G, p.B) for p in pixels ] 
for s,f in zip(shapes,formulas) : s.Cells("FillForegnd").Formula = f 
