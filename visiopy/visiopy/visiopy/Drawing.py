from __future__ import division

class Point(object):
    
    def __init__( self , x, y) :
        self.X = x
        self.Y = y

    def Clone(self) :
        return Point(self.X,self.Y)

    def __repr__(self) :
        return "(%s,%s)" % ( self.X, self.Y)

class Rectangle(object):
    
    def __init__( self , left, bottom, right, top) :
        self.Left = left
        self.Bottom = bottom
        self.Right = right
        self.Top = top

    @property
    def Size(self):
        return Size( self.Right - self.Left, self.Top- self.Bottom)

    @property
    def CenterPoint(self):
        x = self.Left + ((self.Right - self.Left)/2.0)
        y = self.Bottom+ ((self.Top- self.Bottom)/2.0)
        return Point( x, y)

    def Clone(self) :
        return Rectangle(self.Left, self.Bottom, self.Right ,self.Top )

    def __repr__(self) :
        return "(%s,%s,%s,%s)" % ( self.Left, self.Bottom, self.Right, self.Top )

class Size(object):
    
    def __init__( self , w, h) :
        self.Width = w
        self.Height = h

    def Clone(self) :
        return Point(self.Width,self.Height)

    def __repr__(self) :
        return "(%s,%s)" % ( self.Width, self.Height)
