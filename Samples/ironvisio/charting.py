
import math
from geometry import *

class VerticalBarChart :

    def __init__( self) : 
        self.DataPoints = []
        self.Categories = []
        self.MaxHeight = 3.0
        self.BarWidth=1.0
        self.HorizontalDistance=0.5
        self.CategoryHeight = 0.5
        self.CategoryDistance=0.0
        self.Origin = Point(0,0)

    def Draw(self, page) :

        # Calculate Geometry
        numpoints = len(self.DataPoints)
        heights = normalize_to( (p.Value for p in self.DataPoints), self.MaxHeight)
        top_row_origin = self.Origin.AddSize( Size(0, self.CategoryDistance+self.CategoryHeight) )
        top_row_rects = get_rects_horiz_vary_heights( top_row_origin, self.BarWidth, heights, self.HorizontalDistance, numpoints )
        bottom_row_rects = get_rects_horiz( self.Origin , Size(self.BarWidth,self.CategoryHeight), self.HorizontalDistance, numpoints )

        # draw bars
        barshapes = drawrects( page, top_row_rects )
        settext( barshapes, [ p.Label for p in self.DataPoints ] )

        # draw category textboxes
        catshapes = drawrects( page, bottom_row_rects )
        settext( catshapes, self.Categories )


class CircleChart :

    def __init__( self) : 
        self.DataPoints = []
        self.Categories = []
        self.MaxHeight = 1.0
        self.HorizontalDistance=0.5
        self.CategoryHeight = 0.5
        self.CategoryDistance=0.0
        self.Origin = Point(0,0)

    def Draw(self, page) :
        # Calculate Geometry
        numpoints = len(self.DataPoints)
                
        normalized_values = normalize( (p.Value for p in self.DataPoints) )
        radii = [ math.sqrt(v/math.pi) for v in normalized_values]
        radii = normalize_to( radii, self.MaxHeight/2.0 )

        top_row_origin = self.Origin.AddSize( Size(0, self.CategoryDistance+self.CategoryHeight) )
        top_row_rects = get_rects_horiz( top_row_origin, Size(self.MaxHeight, self.MaxHeight), self.HorizontalDistance, numpoints )
        centers = [ r.Center for r in top_row_rects ]
        circlerects = [ Rectangle.FromPointAndRadius(c,r) for (c,r) in zip(centers,radii) ]
        bottom_row_rects = get_rects_horiz( self.Origin , Size(self.MaxHeight,self.CategoryHeight), self.HorizontalDistance, numpoints )

        # draw circle
        circleshapes = drawovals(page, circlerects)
        settext( circleshapes, [p.Label for p in self.DataPoints] )

        # draw category textboxes
        catshapes = drawrects( page, bottom_row_rects )
        settext( catshapes, self.Categories )

def drawrects( page, rects ) :
    shapes = []
    for r in rects:
        shape = page.DrawRectangle(r.X0, r.Y0, r.X1, r.Y1)
        shapes.append(shape)
    return shapes

def drawovals( page, rects ) :
    shapes = []
    for r in rects:
        shape = page.DrawOval(r.X0, r.Y0, r.X1, r.Y1)
        shapes.append(shape)
    return shapes

def settext( shapes, texts) :
    for (shape,text) in zip(shapes,texts) :
        shape.Text = text

def normalize( seq ) :
    items = [v for v in seq]
    m = max( items )
    return [ float(v)/m for v in items ]

def normalize_to( seq , s) :
    items = [v for v in seq]
    m = max( items )
    return [ float(v)/m*s for v in items ]

