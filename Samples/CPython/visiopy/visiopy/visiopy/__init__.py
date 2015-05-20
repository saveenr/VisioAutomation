import array
import sys 
import win32com.client 
win32com.client.gencache.EnsureDispatch("Visio.Application") 

from ShapeSheet import *
from DOM import *
from Drawing import *

if (__name__=='__main__') :
    print "Visiopy Cannot be run as a script."
else :
    pass