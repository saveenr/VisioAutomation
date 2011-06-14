using VisioInterop;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;

public static partial class VS2010_CSharp_Samples
{
    public static void Page_GetFormulas(IVisio.Document doc)
    {
        var page = VisioInterop.Util.CreateStandardPage(doc, "PGF");
        var shape = VisioInterop.Util.CreateStandardShape(page);
        var request = new []
        {
              new
                  {
                      ID=shape.ID16, 
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormWidth
                  },                        
              new
                  {
                      ID=shape.ID16, 
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormWidth
                  }                        
        };

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        var SID_SRCStream = new short[request.Length*4];
        for (int i = 0; i < request.Length; i++)
        {
            SID_SRCStream.Set4(i, request[i].ID, request[i].Section, request[i].Row, request[i].Cell);
        }

        // EXECUTE THE REQUEST
        System.Array formulas_sa;
        page.GetFormulasU(SID_SRCStream, out formulas_sa);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        var formulas = new string[formulas_sa.Length];
        formulas_sa.CopyTo(formulas, 0);

        // DISPLAY THE INFORMATION
        shape.Text = string.Format("Formulas={0},{1}", formulas[0], formulas[1]);
    }
}