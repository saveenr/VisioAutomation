using IVisio = Microsoft.Office.Interop.Visio;

public static partial class CSharpSamples
{
    public static void Shape_GetResults(IVisio.Document doc)
    {
        var pages = doc.Pages;
        var page = pages.Add();
        page.NameU = "SGR";

        var shape = page.DrawRectangle(1, 1, 4, 3);
        shape.get_CellsU("Width").Formula = "=(1.0+2.5)";
        shape.get_CellsU("Height").Formula = "=(0.0+1.5)";

        // BUILD UP THE REQUEST
        short flags = (short)IVisio.VisGetSetArgs.visGetFloats;
        var items = new[]
        {
            new {   section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormWidth,
                    unitcode= (short) IVisio.VisUnitCodes.visNoCast},
            
            new {   section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormHeight,
                    unitcode= (short) IVisio.VisUnitCodes.visNoCast}
        };

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        short[] SRCStream = new short[items.Length * 3];
        object[] unitcodes = new object[items.Length];
        for (int i = 0; i < items.Length; i++)
        {
            SRCStream[i * 3 + 0] = items[i].section;
            SRCStream[i * 3 + 1] = items[i].row;
            SRCStream[i * 3 + 2] = items[i].cell;
            unitcodes[i] = items[i].unitcode;
        }

        // EXECUTE THE REQUEST
        System.Array results_sa;
        System.Array SRCStream_sa = (System.Array) SRCStream;
        System.Array unitcodes_sa = (System.Array) unitcodes;
        shape.GetResults(ref SRCStream_sa, flags, ref unitcodes_sa, out results_sa);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        object[] results_objects = (object[])results_sa;
        double [] results_doubles = new double[results_objects.Length];
        results_objects.CopyTo(results_doubles, 0);

        // DISPLAY THE INFORMATION
		shape.Text = string.Format("Results={0},{1}", results_doubles[0], results_doubles[1]);
    }
}
