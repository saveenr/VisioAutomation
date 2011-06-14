using System.Collections.Generic;
using VisioInterop;
using IVisio = Microsoft.Office.Interop.Visio;

public static partial class VS2010_CSharp_Samples
{
    public static void Page_SetFormulas(IVisio.Document doc)
    {
        var page = VisioInterop.Util.CreateStandardPage(doc, "PSF");
        var shape = VisioInterop.Util.CreateStandardShape(page);
        var request = new[]
        {
              new
                  {
                      ID=shape.ID16, 
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormWidth,
                      Formula="2.0"
                  },                        
              new
                  {
                      ID=shape.ID16, 
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormHeight,
                      Formula="3.0"
                  }                        
        };


        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS

        var SID_SRCStream = new short[request.Length*4];
        var formulas_objects = new object[request.Length];
        for (int i = 0; i < request.Length; i++)
        {
            SID_SRCStream[(i * 4) + 0] = request[i].ID;
            SID_SRCStream[(i * 4) + 1] = request[i].Section;
            SID_SRCStream[(i * 4) + 2] = request[i].Row;
            SID_SRCStream[(i * 4) + 3] = request[i].Cell;
            formulas_objects[i] = request[i].Formula;
        }

        // EXECUTE THE REQUEST
        short flags = (short)(IVisio.VisGetSetArgs.visSetBlastGuards | IVisio.VisGetSetArgs.visSetUniversalSyntax);
        int count = page.SetFormulas(SID_SRCStream, formulas_objects, flags);

        // DISPLAY THE INFORMATION
        shape.Text = "SetFormulas";
    }
}