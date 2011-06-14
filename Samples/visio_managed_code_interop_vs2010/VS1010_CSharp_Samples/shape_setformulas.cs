using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioInterop;

public static partial class VS2010_CSharp_Samples
{
    public static void Shape_SetFormulas(IVisio.Document doc)
    {
        var page = VisioInterop.Util.CreateStandardPage(doc, "SSF");
        var shape = VisioInterop.Util.CreateStandardShape(page);
        var request = VisioInterop.Util.Create_ShapeSetFormulas_Request();

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS

        var SRCStream = new short[request.Length * 3];
        var formulas_objects = new object[request.Length];
        for (int i = 0; i < request.Length; i++)
        {
            SRCStream.Set3(i, request[i].CellSRC.SectionIndex, request[i].CellSRC.RowIndex, request[i].CellSRC.CellIndex);
            formulas_objects[i] = request[i].Formula;
        }

        // EXECUTE THE REQUEST
        short flags = (short)(IVisio.VisGetSetArgs.visSetBlastGuards | IVisio.VisGetSetArgs.visSetUniversalSyntax);
        int count = shape.SetFormulas(SRCStream, formulas_objects, flags);

        // DISPLAY THE INFORMATION
        shape.Text = "SetFormulas";
    }
}