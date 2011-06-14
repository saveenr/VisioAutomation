using System.Collections.Generic;
using VisioInterop;
using IVisio = Microsoft.Office.Interop.Visio;

public static partial class VS2010_CSharp_Samples
{
    public static void Page_SetResults(IVisio.Document doc)
    {
        var page = VisioInterop.Util.CreateStandardPage(doc, "PSR");
        var shape = VisioInterop.Util.CreateStandardShape(page);
        var request = Util.Create_PageSetResults_Request(shape);

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS


        var SID_SRCStream = new short[request.Length * 4];
        var results_objects = new object[request.Length];
        var unitcodes = new object[request.Length];
        for (int i = 0; i < request.Length; i++)
        {
            SID_SRCStream.Set4(i,request[i].ShapeID, request[i].CellSRC.SectionIndex, request[i].CellSRC.RowIndex, request[i].CellSRC.CellIndex);
            results_objects[i] = request[i].Result;
            unitcodes[i] = request[i].UnitCode;
        }

        // EXECUTE THE REQUEST
        short flags = 0;
        int count = page.SetResults(SID_SRCStream, unitcodes, results_objects, flags);

        // DISPLAY THE INFORMATION
        shape.Text = "SetResults";
    }
}