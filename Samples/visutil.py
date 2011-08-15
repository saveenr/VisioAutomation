import clr 
import System 

clr.AddReference("Microsoft.Office.Interop.Visio") 
import Microsoft.Office.Interop.Visio 
IVisio = Microsoft.Office.Interop.Visio 

def load_stencil( visapp, filename ) :
        stencilflags = IVisio.VisOpenSaveArgs.visOpenRO | IVisio.VisOpenSaveArgs.visOpenDocked 
        stencildoc = visapp.Documents.OpenEx(filename, stencilflags) 
        return stencildoc

def get_ref_short_array() :
    ref_array = clr.Reference[System.Array[System.Int16]](System.Array[System.Int16]([])) 
    return ref_array

def dropmany( page, masters, xys ) :
        masters_obj_arr = System.Array[object]( masters ) 
        xys = System.Array[System.Double]( xys ) 
        out_ids = get_ref_short_array() 
        page.DropManyU( masters_obj_arr, xys, out_ids ) 
        return out_ids

def set_formulas( page, items ) :
    num_items = len(items)
    SID_SRCStream = System.Array[System.Int16][num_items* 4];
    formulas_objects = System.Array[System.Int16][num_items];
    for i in xrange(num_items):
        SID_SRCStream[i * 4 + 0] = items[i][0];
        SID_SRCStream[i * 4 + 1] = items[i][1];
        SID_SRCStream[i * 4 + 2] = items[i][2];
        SID_SRCStream[i * 4 + 3] = items[i][3];
        formulas_objects[i] = items[i][4];

def test() :


        visapp = IVisio.ApplicationClass() 
        doc = visapp.Documents.Add("") 
        page = visapp.ActivePage 

        stencildoc = load_stencil( visapp, "basic_u.vss" )
        rectangle = stencildoc.Masters[ "Rectangle" ]
        masters = [ rectangle ]
        xys = [ 1.0, 2.0 , 5.0, 6.0 ]
        dropmany( page, masters, xys )

        ids = draw_grid(page, 0,0, [1,1.5,2] , [0.5,0.5])

        print ids

        items = [ (ids[0],1,2,3,"F") ]
        set_formulas( page, items )

def draw_grid( page, left, top, widths, heights ) :

    lefts = []
    cur=left
    for width in widths :
        lefts.append(cur)
        cur += width

    tops = []
    cur = top
    for height in heights :
        tops.append(cur)
        cur -= height

    print lefts
    print tops

    xys = [ ]
    for row,ctop in enumerate(tops) :
        for col,cleft in enumerate(lefts) :
            cright = cleft + widths[col]
            cbottom = ctop - heights[row]
            xys.extend( [(cleft + cright)/2.0, (cbottom+ctop)/2.0] )

    stencildoc = load_stencil( page.Application, "basic_u.vss" )
    master = stencildoc.Masters[ "Rectangle" ]
    masters = [ master ]
    ids = dropmany( page, masters, xys )
    return ids
    
        
if __name__ == "__main__" :
        test()
        
