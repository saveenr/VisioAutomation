using IVisio = Microsoft.Office.Interop.Visio;

public static partial class CSharpSamples
{
    public static void Page_GetResults(IVisio.Document doc)
    {
        var pages = doc.Pages;
        var page = pages.Add();
        page.NameU = "PGR";

        var shape = page.DrawRectangle(1, 1, 4, 3);
        shape.get_CellsU("Width").Formula = "=(1.0+2.5)";
        shape.get_CellsU("Height").Formula = "=(0.0+1.5)";

        // BUILD UP THE REQUEST
        var flags = (short)IVisio.VisGetSetArgs.visGetFloats;
        var items = new[]
        {
            new {   shapeid = (short) shape.ID,
                    section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormWidth,
                    unitcode= (short) IVisio.VisUnitCodes.visNoCast },
            
            new {   shapeid = (short) shape.ID,
                    section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormHeight,
                    unitcode= (short) IVisio.VisUnitCodes.visNoCast }
        };

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        short[] SID_SRCStream = new short[items.Length * 4];
        object[] unitcodes = new object[items.Length];
        for (int i = 0; i < items.Length; i++)
        {
            SID_SRCStream[i * 4 + 0] = items[i].shapeid;
            SID_SRCStream[i * 4 + 1] = items[i].section;
            SID_SRCStream[i * 4 + 2] = items[i].row;
            SID_SRCStream[i * 4 + 3] = items[i].cell;
            unitcodes[i] = items[i].unitcode;
        }

        // EXECUTE THE REQUEST
        System.Array results_sa;
		System.Array SID_SRCStream_sa = (System.Array) SID_SRCStream;
		System.Array unitcodes_sa = (System.Array) unitcodes;
		page.GetResults(ref SID_SRCStream_sa, flags, ref unitcodes_sa, out results_sa);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        object[] results_objects = (object[])results_sa;
        double [] results_doubles = new double[results_objects.Length];
        results_objects.CopyTo(results_doubles, 0);

        // DISPLAY THE INFORMATION
		shape.Text = string.Format("Results={0},{1}", results_doubles[0], results_doubles[1]);
    }
}
