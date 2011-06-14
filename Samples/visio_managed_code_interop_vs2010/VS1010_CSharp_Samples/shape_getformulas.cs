using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioInterop;
public static partial class VS2010_CSharp_Samples
{
    public static void Shape_GetFormulas(IVisio.Document doc)
    {
        var page = VisioInterop.Util.CreateStandardPage(doc, "SGF");
        var shape = VisioInterop.Util.CreateStandardShape(page);
        var request = new[]
        {
              new
                  {
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormWidth
                  },                        
              new
                  {
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormHeight,
                  }                        
        };

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        var SRCStream = new short[request.Length * 3];
        for (int i = 0; i < request.Length; i++)
        {
            SRCStream[(i * 3) + 0] = request[i].Section;
            SRCStream[(i * 3) + 1] = request[i].Row;
            SRCStream[(i * 3) + 2] = request[i].Cell;
        }

        // EXECUTE THE REQUEST
        System.Array formulas_sa;
        shape.GetFormulasU(SRCStream, out formulas_sa);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        var formulas = new string[formulas_sa.Length];
        formulas_sa.CopyTo(formulas, 0);

        // DISPLAY THE INFORMATION
        shape.Text = string.Format("Formulas={0},{1}", formulas[0], formulas[1]);
    }
}