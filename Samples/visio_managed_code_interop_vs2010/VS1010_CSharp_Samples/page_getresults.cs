using System.Collections.Generic;
using VisioInterop;
using IVisio = Microsoft.Office.Interop.Visio;

public static partial class VS2010_CSharp_Samples
{
    public static void Page_GetResults(IVisio.Document doc)
    {
        var page = VisioInterop.Util.CreateStandardPage(doc, "PGR");
        var shape = VisioInterop.Util.CreateStandardShape(page);
        var request = Util.Create_PageGetResults_Request(shape);

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        var SID_SRCStream = new short[request.Length * 4];
        var unitcodes = new object[request.Length];
        for (int i = 0; i < request.Length; i++)
        {
            SID_SRCStream.Set4(i, request[i].ShapeID, request[i].CellSRC.Section, request[i].CellSRC.Row, request[i].CellSRC.Cell);
            unitcodes[i] = request[i].UnitCode;
        }

        // EXECUTE THE REQUEST
        var flags = (short)IVisio.VisGetSetArgs.visGetFloats;
        System.Array results_sa;
        page.GetResults(SID_SRCStream, flags, unitcodes, out results_sa);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        var results_doubles = new double[results_sa.Length];
        results_sa.CopyTo(results_doubles, 0);

        // DISPLAY THE INFORMATION
        shape.Text = string.Format("Results={0},{1}", results_doubles[0], results_doubles[1]);
    }
}