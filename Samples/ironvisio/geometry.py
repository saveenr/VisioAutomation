
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

