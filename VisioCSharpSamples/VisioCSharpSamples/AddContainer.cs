using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioCSharpSamples
{

    public static partial class Samples
    {
        public static void AddContainer()
        {
            var app = new IVisio.ApplicationClass();

            var docs = app.Documents;
            var doc = docs.Add("");
            var page = app.ActivePage;
            short flags = (short)
                (IVisio.VisOpenSaveArgs.visOpenDocked |
                IVisio.VisOpenSaveArgs.visOpenRO);

            var basic_stencil = docs.OpenEx("basic_u.vss", flags);
            var basic_masters = basic_stencil.Masters;
            var rounded_rect_master = basic_masters["Rounded Rectangle"];

            double x = 2.0;
            double y = 3.0;

            var shape1 = page.Drop(rounded_rect_master, x, y);
            var shape2 = page.Drop(rounded_rect_master, x + 3.0, y + 1.0);

            var stenciltype = IVisio.VisBuiltInStencilTypes.visBuiltInStencilContainers;

            var measurementsys = IVisio.VisMeasurementSystem.visMSDefault;
            short containeropenflags = (short)IVisio.VisOpenSaveArgs.visOpenHidden;
            string containerstencil_filename = app.GetBuiltInStencilFile(stenciltype, measurementsys);
            var container_stencil = docs.OpenEx(containerstencil_filename, containeropenflags);
            var containermasters = container_stencil.Masters;
            var container = containermasters["Container 1"];

            var activewindow = app.ActiveWindow;
            short selectargs = (short)IVisio.VisSelectArgs.visSelect;
            activewindow.Select(shape1, selectargs);
            activewindow.Select(shape2, selectargs);

            page.DropContainer(container, activewindow.Selection);

        }
    }
}