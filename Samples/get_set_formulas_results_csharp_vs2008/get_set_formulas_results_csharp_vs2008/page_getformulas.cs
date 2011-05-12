using IVisio = Microsoft.Office.Interop.Visio;

public static partial class CSharpSamples
{
    public static void Page_GetFormulas(IVisio.Document doc)
    {
        var pages = doc.Pages;
        var page = pages.Add();
        page.NameU = "PGF";

        var shape = page.DrawRectangle(1, 1, 4, 3);
        shape.get_CellsU("Width").Formula = "=(1.0+2.5)";
        shape.get_CellsU("Height").Formula = "=(0.0+1.5)";

        // BUILD UP THE REQUEST
        var items = new[]
        {
            new {   shapeid = (short) shape.ID,
                    section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormWidth },
            
            new {   shapeid = (short) shape.ID,
                    section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormHeight }
        };

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        short[] SID_SRCStream = new short[items.Length * 4];
        for (int i = 0; i < items.Length; i++)
        {
            SID_SRCStream[i * 4 + 0] = items[i].shapeid;
            SID_SRCStream[i * 4 + 1] = items[i].section;
            SID_SRCStream[i * 4 + 2] = items[i].row;
            SID_SRCStream[i * 4 + 3] = items[i].cell;
        }

        // EXECUTE THE REQUEST
        System.Array formulas_sa;
        System.Array SID_SRCStream_sa = (System.Array) SID_SRCStream;
        page.GetFormulasU(ref SID_SRCStream_sa, out formulas_sa);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        object[] formulas_objects = (object[])formulas_sa;
        string [] formulas = new string[formulas_objects.Length];
        formulas_objects.CopyTo(formulas, 0);

        // DISPLAY THE INFORMATION
        shape.Text = string.Format("Formulas={0},{1}", formulas[0], formulas[1]);
    }
}
