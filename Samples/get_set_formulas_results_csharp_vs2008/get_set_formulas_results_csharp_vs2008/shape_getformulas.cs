using IVisio = Microsoft.Office.Interop.Visio;

public static partial class CSharpSamples
{
    public static void Shape_GetFormulas(IVisio.Document doc)
    {
        var pages = doc.Pages;
        var page = pages.Add();
        page.NameU = "SGF";

        var shape = page.DrawRectangle(1, 1, 4, 3);
        shape.get_CellsU("Width").Formula = "=(1.0+2.5)";
        shape.get_CellsU("Height").Formula = "=(0.0+1.5)";

        // BUILD UP THE REQUEST
        var items = new[]
        {
            new {   section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormWidth },
            
            new {   section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormHeight }
        };

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        short[] SRCStream = new short[items.Length * 3];
        for (int i = 0; i < items.Length; i++)
        {
            SRCStream[i * 3 + 0] = items[i].section;
            SRCStream[i * 3 + 1] = items[i].row;
            SRCStream[i * 3 + 2] = items[i].cell;
        }

        // EXECUTE THE REQUEST
        System.Array formulas_sa;
        System.Array SRCStream_sa = (System.Array) SRCStream;
        shape.GetFormulasU(ref SRCStream_sa, out formulas_sa);

        // MAP OUTPUT BACK TO SOMETHING USEFUL 
        object [] formulas_objects = (object[])formulas_sa;
        string [] formulas = new string[formulas_objects.Length];
        formulas_objects.CopyTo(formulas, 0);

        // DISPLAY THE INFORMATION
        shape.Text = string.Format("Formulas={0},{1}", formulas[0], formulas[1]);
    }
}
