import clr 
import System 

clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 
IVisio = Microsoft.Office.Interop.Visio 

from Records import *
import Shape_GetFormulas

def get_new_system_array(T,length) :
    array = System.Array.CreateInstance(T, length)
    return array

def get_ref_to_system_array(T,array) :
    ref = clr.Reference[System.Array[T]](array) 
    return ref

def get_outref_to_system_array(T) :
    ref = clr.Reference[System.Array[T]](System.Array[T]([])) 
    return ref
        
        
