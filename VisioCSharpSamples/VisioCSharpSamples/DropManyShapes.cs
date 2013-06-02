using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioCSharpSamples
{

    public static partial class Samples
    {
        public static void DropManyShapes()
        {
            var app = new IVisio.ApplicationClass();
            var docs = app.Documents;
            var doc = docs.Add("");
            var page = app.ActivePage;
            short flags = (short)
                ( IVisio.VisOpenSaveArgs.visOpenDocked | 
                IVisio.VisOpenSaveArgs.visOpenRO);
            
            var basic_stencil = docs.OpenEx("basic_u.vss", flags);
            var basic_masters = basic_stencil.Masters;
            var rounded_rect_master = basic_masters["Rounded Rectangle"];

            // three masters to drop
            var masters = new object[] { rounded_rect_master, rounded_rect_master, rounded_rect_master };

            // three coordinates to make the drop
            var xy_array = new double[] { 1, 2, 3, 5, 6, 2 };

            // perform the drop of all shapes at once
            // and get the shape ids back out
            System.Array shape_ids_sa;
            page.DropManyU(masters, xy_array, out shape_ids_sa);
            var shape_ids = (short[])shape_ids_sa;
        }
    }
}