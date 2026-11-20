using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Shapes;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page geometry.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class GeometrySamples
    {
        public static void Sample_01(IVisio.Document doc, double a)
        {
            // <sample id="geometry#01" doc="geometry.md" line="13">
            var page  = doc.Pages.Add();
            var shape = page.DrawRectangle(0, 0, 4, 4);

            // Replace the rectangle's default geometry with a custom triangle
            VisioAutomation.Shapes.GeometryHelper.Delete(shape);

            var geom = new VisioAutomation.Shapes.GeometrySection();
            geom.AddMoveTo(0, 0);
            geom.AddLineTo(4, 0);
            geom.AddLineTo(2, 3);
            geom.AddLineTo(0, 0);

            geom.Render(shape);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape)
        {
            // <sample id="geometry#02" doc="geometry.md" line="43">
            var geom = new VisioAutomation.Shapes.GeometrySection();
            geom.NoLine = 1;     // fill only, no outline
            geom.AddMoveTo(0, 0);
            geom.AddLineTo(2, 0);
            geom.AddLineTo(2, 2);
            geom.AddLineTo(0, 0);
            geom.Render(shape);
            // </sample>
        }

        public static void Sample_03(double x, double y, double a)
        {
            // <sample id="geometry#03" doc="geometry.md" line="57">
            GeometryRow.CreateMoveTo(x, y);
            GeometryRow.CreateLineTo(x, y);
            GeometryRow.CreateArcTo(x, y, a);          // a = arc bulge
            // </sample>
        }

        public static void Sample_04(IVisio.Shape shape)
        {
            // <sample id="geometry#04" doc="geometry.md" line="71">
            VisioAutomation.Shapes.GeometryHelper.Delete(shape);
            // </sample>
        }
    }
}
