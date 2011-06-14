using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioInterop;
public static partial class VS2010_CSharp_Samples
{
	public static void Shape_SetResults(IVisio.Document doc)
	{
        var page = VisioInterop.Util.CreateStandardPage(doc, "SSR");
        var shape = VisioInterop.Util.CreateStandardShape(page);
        var request = VisioInterop.Util.Create_ShapeSetResults_Request();

		// MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS

        var SRCStream = new short[request.Length * 3];
        var results_objects = new object[request.Length];
        var unitcodes = new object[request.Length];
        for (int i = 0; i < request.Length; i++)
        {
            SRCStream.Set3(i,request[i].CellSRC.Section, request[i].CellSRC.Row, request[i].CellSRC.Cell);
            results_objects[i] = request[i].Result;
            unitcodes[i] = request[i].UnitCode;
        }

		// EXECUTE THE REQUEST
        short flags = 0;
        int count = shape.SetResults(SRCStream, unitcodes, results_objects, flags);

        // DISPLAY THE INFORMATION
		shape.Text = "SetResults";
	}
}