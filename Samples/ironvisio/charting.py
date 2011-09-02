
import math

class DataPoint :

    def __init__( self, value, label = None ) : 
        self.Value = value
        
        if (label==None) : 
            self.Label = str(value)
        else: 
            self.Label = label

class Size:

    def __init__( self, w, h) : 
        self.Width = w
        self.Height = h

class Point:

    def __init__( self, x, y) : 
        self.X = x
        self.Y = y

    def AddSize( self, s ) :
        return Point( self.X + s.Width, self.Y + s.Height )

class Rectangle :

    def __init__( self, x0, y0, x1, y1) : 
        self.X0 = x0
        self.Y0 = y0
        self.X1 = x1
        self.Y1 = y1

    @property
    def Width(self):
        return self.X1-self.X0

    @property
    def LowerLeft(self):
        return Point(self.X0,self.Y0)

    @property
    def UpperRight(self):
        return Point(self.X1,self.Y1)

    @property
    def Height(self):
        return self.Y1-self.Y0

    def GetFloats(self):
        return (self.X0, self.Y0, self.X1, self.Y1)

    @property
    def Center(self):
        return Point( (self.X0+self.X1)/2.0, (self.Y0+self.Y1)/2.0)

    @staticmethod
    def FromPointAndSize(p,s):
        return Rectangle(p.X, p.Y, p.X+s.Width, p.Y+s.Height)    

    @staticmethod
    def FromPointAndRadius(c,r):
        return Rectangle(c.X-r,c.Y-r,c.X+r,c.Y+r)    

class VerticalBarChart :

    def __init__( self) : 
        self.DataPoints = []
        self.Categories = []
        self.MaxHeight = 3.0
        self.BarWidth=1.0
        self.BarDistance=0.5
        self.CategoryHeight = 0.5
        self.CategoryDistance=0.0
        self.Origin = Point(0,0)

    def Draw(self, page) :
        # Calculate Geometry
        numpoints = len(self.DataPoints)
        heights = normalize_to( (p.Value for p in self.DataPoints), self.MaxHeight)
        bar_origin = self.Origin.AddSize( Size(0, self.CategoryDistance+self.CategoryHeight) )
        bar_rects = get_rects_horiz_vary_heights( bar_origin, self.BarWidth, heights, self.BarDistance, numpoints )
        cat_rects = get_rects_horiz( self.Origin , Size(self.BarWidth,self.CategoryHeight), self.BarDistance, numpoints )

        # draw bars
        barshapes = drawrects( page, bar_rects )
        settext( barshapes, [ p.Label for p in self.DataPoints ] )

        # draw category textboxes
        catshapes = drawrects( page, cat_rects )
        settext( catshapes, self.Categories )


class CircleChart :

    def __init__( self) : 
        self.DataPoints = []
        self.Categories = []
        self.MaxRadius= 0.5
        self.CircleDistance=0.5
        self.CategoryHeight = 0.5
        self.CategoryDistance=0.0
        self.Origin = Point(0,0)

    def Draw(self, page) :
        # Calculate Geometry
        numpoints = len(self.DataPoints)
                
        normalized_values = normalize( (p.Value for p in self.DataPoints) )
        radii = [ math.sqrt(v/math.pi) for v in normalized_values]
        radii = normalize_to( radii, self.MaxRadius )

        bar_origin = self.Origin.AddSize( Size(0, self.CategoryDistance+self.CategoryHeight) )
        bar_rects = get_rects_horiz( bar_origin, Size(self.MaxRadius*2, self.MaxRadius*2), self.CircleDistance, numpoints )
        centers = [ r.Center for r in bar_rects ]
        circlerects = [ Rectangle.FromPointAndRadius(c,r) for (c,r) in zip(centers,radii) ]
        cat_rects = get_rects_horiz( self.Origin , Size(2*self.MaxRadius,self.CategoryHeight), self.CircleDistance, numpoints )

        # draw circle
        circleshapes = drawovals(page, circlerects)
        settext( circleshapes, [p.Label for p in self.DataPoints] )

        # draw category textboxes
        catshapes = drawrects( page, cat_rects )
        settext( catshapes, self.Categories )

def get_points_horiz( origin, sep, count ) :
    points = [ Point(origin.X + i*sep,origin.Y) for i in xrange(count)]
    return points

def get_rects_horiz( origin, size, sep, count ) :
    skip = size.Width + sep
    lowerlefts = get_points_horiz( origin, skip, count )
    rects = [ Rectangle.FromPointAndSize( ll , size ) for ll in lowerlefts ]
    return rects

def get_rects_horiz_vary_heights( origin, width, heights, sep, count , ) :
    skip = width + sep
    lowerlefts = get_points_horiz( origin, skip, count )
    rects = [ Rectangle.FromPointAndSize( ll , Size(width,h) ) for (ll,h) in zip(lowerlefts,heights) ]
    return rects

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

