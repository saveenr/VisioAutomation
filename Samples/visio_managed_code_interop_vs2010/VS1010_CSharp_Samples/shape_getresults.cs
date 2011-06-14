using IVisio = Microsoft.Office.Interop.Visio;

public static partial class VS2010_CSharp_Samples
{
    public static void Shape_GetResults(IVisio.Document doc)
    {
        var page = VisioInterop.Util.CreateStandardPage(doc, "SGR");
        var shape = VisioInterop.Util.CreateStandardShape(page);

        // CREATE REQUEST
        var request = new[]
        {
              new
                  {
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormWidth,
                      UnitCode=(short) IVisio.VisUnitCodes.visNoCast
                  },                        
              new
                  {
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormHeight,
                      UnitCode=(short) IVisio.VisUnitCodes.visNoCast
                  }                        
        };

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        var SRCStream = new short[request.Length * 3];
        var unitcodes = new object[request.Length];
        for (int i = 0; i < request.Length; i++)
        {
            SRCStream[(i * 3) + 0] = request[i].Section;
            SRCStream[(i * 3) + 1] = request[i].Row;
            SRCStream[(i * 3) + 2] = request[i].Cell;
            unitcodes[i] = request[i].UnitCode;
        }

        // EXECUTE THE REQUEST
        short flags = (short)IVisio.VisGetSetArgs.visGetFloats;
        System.Array results_sa = null;
        shape.GetResults(SRCStream, flags, unitcodes, out results_sa);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        var results_doubles = new double[results_sa.Length];
        results_sa.CopyTo(results_doubles, 0);

        // DISPLAY THE INFORMATION
        shape.Text = string.Format("Results={0},{1}", results_doubles[0], results_doubles[1]);
    }
}