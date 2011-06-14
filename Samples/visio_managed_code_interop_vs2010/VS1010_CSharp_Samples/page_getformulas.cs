using IVisio = Microsoft.Office.Interop.Visio;

public static partial class VS2010_CSharp_Samples
{
    public static void Page_GetFormulas(IVisio.Document doc)
    {
        var page = Util.CreateStandardPage(doc, "PGF");
        var shape = Util.CreateStandardShape(page);

        // CREATE REQUEST
        var request = new []
        {
              new
                  {
                      ID=(short)shape.ID16, 
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormWidth
                  },                        
              new
                  {
                      ID=(short)shape.ID16, 
                      Section = (short)IVisio.VisSectionIndices.visSectionObject, 
                      Row=(short)IVisio.VisRowIndices.visRowXFormOut, 
                      Cell=(short)IVisio.VisCellIndices.visXFormHeight,
                  }                        
        };

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        var SID_SRCStream = new short[request.Length*4];
        for (int i = 0; i < request.Length; i++)
        {
            SID_SRCStream[(i * 4) + 0] = request[i].ID;
            SID_SRCStream[(i * 4) + 1] = request[i].Section;
            SID_SRCStream[(i * 4) + 2] = request[i].Row;
            SID_SRCStream[(i * 4) + 3] = request[i].Cell;
        }

        // EXECUTE THE REQUEST
        System.Array formulas_sa = null;
        page.GetFormulasU(SID_SRCStream, out formulas_sa);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        var formulas = new string[formulas_sa.Length];
        formulas_sa.CopyTo(formulas, 0);

        // DISPLAY THE INFORMATION
        shape.Text = string.Format("Formulas={0},{1}", formulas[0], formulas[1]);
    }
}