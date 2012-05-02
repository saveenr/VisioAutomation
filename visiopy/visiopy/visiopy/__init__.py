import array
import sys 
import win32com.client 
win32com.client.gencache.EnsureDispatch("Visio.Application") 

from ShapeSheet import *
from DOM import *
from Drawing import *

def openstencil(docs, stencilname) :
    stencildocflags = win32com.client.constants.visOpenRO | win32com.client.constants.visOpenDocked 
    stencildoc = docs.OpenEx(stencilname , stencildocflags )
    return stencildoc

if (__name__=='__main__') :
    pass
else :
    pass