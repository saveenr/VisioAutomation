using IVisio = Microsoft.Office.Interop.Visio;

public static partial class CSharpSamples
{
    public static void Page_SetResults(IVisio.Document doc)
    {
        var pages = doc.Pages;
        var page = pages.Add();
        page.NameU = "PSR";

        var shape = page.DrawRectangle(1, 1, 4, 3);
        shape.get_CellsU("Width").Formula = "=(1.0+2.5)";
        shape.get_CellsU("Height").Formula = "=(0.0+1.5)";

        // BUILD UP THE REQUEST
        short flags = 0;
        var items = new[]
        {
            new {   shapeid = (short) shape.ID,
                    section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormWidth,
                    result = 8.0 ,
                    unitcode = (short) IVisio.VisUnitCodes.visNoCast },    

            new {   shapeid = (short) shape.ID,
                    section = (short) IVisio.VisSectionIndices.visSectionObject, 
                    row     = (short) IVisio.VisRowIndices.visRowXFormOut, 
                    cell    = (short) IVisio.VisCellIndices.visXFormHeight,
                    result = 1.0 ,
                    unitcode = (short) IVisio.VisUnitCodes.visNoCast }
        };

        // MAP THE REQUEST TO THE STRUCTURES VISIO EXPECTS
        short[] SID_SRCStream = new short[items.Length * 4];
        object[] results_objects = new object[items.Length];
        object[] unitcodes = new object[items.Length];
        for (int i = 0; i < items.Length; i++)
        {
            SID_SRCStream[i * 4 + 0] = items[i].shapeid;
            SID_SRCStream[i * 4 + 1] = items[i].section;
            SID_SRCStream[i * 4 + 2] = items[i].row;
            SID_SRCStream[i * 4 + 3] = items[i].cell;
            results_objects[i] = items[i].result;
            unitcodes[i] = items[i].unitcode;
        }

        // EXECUTE THE REQUEST
		System.Array SID_SRCStream_sa = (System.Array) SID_SRCStream;
		System.Array results_objects_sa = (System.Array) results_objects;
		System.Array unitcodes_sa = (System.Array) unitcodes;
        int count = page.SetResults(ref SID_SRCStream_sa, ref unitcodes_sa, ref results_objects_sa, flags);

        shape.Text = "SetResults";
    }
}
