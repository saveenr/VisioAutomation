from __future__ import division

class Point(object):
    
    def __init__( self , x, y) :
        self.X = x
        self.Y = y

    def Clone() :
        return Point(self.X,self.Y)

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
        return Point( (self.Right - self.Left)/2.0, (self.Top- self.Bottom)/2.0)

    def Clone() :
        return Rectangle(self.Left, self.Bottom, self.Right ,self.Top )

class Size(object):
    
    def __init__( self , w, h) :
        self.Width = w
        self.Height = h

    def Clone() :
        return Point(self.Width,self.Height)
