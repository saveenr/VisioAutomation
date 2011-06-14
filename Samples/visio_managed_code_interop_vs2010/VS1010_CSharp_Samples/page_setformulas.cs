using System.Collections.Generic;
using VisioInterop;
using IVisio = Microsoft.Office.Interop.Visio;

public static partial class VS2010_CSharp_Samples
{
    public static void Page_SetFormulas(IVisio.Document doc)
    {
        var page = VisioInterop.Util.CreateStandardPage(doc, "PSF");
        var shape = VisioInterop.Util.CreateStandardShape(page);
        var request = Util.Create_PageSetFormulas_Request(shape);


        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS

        var SID_SRCStream = new short[request.Length*4];
        var formulas_objects = new object[request.Length];
        for (int i = 0; i < request.Length; i++)
        {
            SID_SRCStream.Set4(i, request[i].ShapeID, request[i].CellSRC.SectionIndex, request[i].CellSRC.RowIndex, request[i].CellSRC.CellIndex);
            formulas_objects[i] = request[i].Formula;
        }

        // EXECUTE THE REQUEST
        short flags = (short)(IVisio.VisGetSetArgs.visSetBlastGuards | IVisio.VisGetSetArgs.visSetUniversalSyntax);
        int count = page.SetFormulas(SID_SRCStream, formulas_objects, flags);

        // DISPLAY THE INFORMATION
        shape.Text = "SetFormulas";
    }
}