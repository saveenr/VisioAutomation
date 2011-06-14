using System.Collections.Generic;
using VisioInterop;
using IVisio = Microsoft.Office.Interop.Visio;

public static partial class VS2010_CSharp_Samples
{
    public static void Page_GetResults(IVisio.Document doc)
    {
        var page = VisioInterop.Util.CreateStandardPage(doc, "PGR");
        var shape = VisioInterop.Util.CreateStandardShape(page);
        var request = new[]
        {
              new
                  {
                      ID=shape.ID16, 
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormWidth,
                      UnitCode=(short) IVisio.VisUnitCodes.visNoCast
                  },                        
              new
                  {
                      ID=shape.ID16, 
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormHeight,
                      UnitCode=(short) IVisio.VisUnitCodes.visNoCast
                  }                        
        };

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        var SID_SRCStream = new short[request.Length * 4];
        var unitcodes = new object[request.Length];
        for (int i = 0; i < request.Length; i++)
        {
            SID_SRCStream[(i * 4) + 0] = request[i].ID;
            SID_SRCStream[(i * 4) + 1] = request[i].Section;
            SID_SRCStream[(i * 4) + 2] = request[i].Row;
            SID_SRCStream[(i * 4) + 3] = request[i].Cell;
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