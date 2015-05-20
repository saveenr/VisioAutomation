using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioCSharpSamples
{

    public static partial class Samples
    {
        public static void LoadStencilDropShape()
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

            double x = 2.0;
            double y = 3.0;
            page.Drop(rounded_rect_master, x, y);
        }
    }
}