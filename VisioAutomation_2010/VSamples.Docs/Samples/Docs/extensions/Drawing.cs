using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page extensions/drawing.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ExtensionsDrawingSamples
    {
        public static void Sample_01()
        {
            // <sample id="extensions/drawing#01" doc="extensions/drawing.md" line="9">

            // </sample>
        }

        public static void Sample_02(IVisio.Page page)
        {
            // <sample id="extensions/drawing#02" doc="extensions/drawing.md" line="27">
            var rect = new VisioAutomation.Core.Rectangle(0, 0, 4, 3);
            var shape = page.DrawRectangle(rect);

            var pts = new[] {
                new VisioAutomation.Core.Point(0, 0),
                new VisioAutomation.Core.Point(2, 1),
                new VisioAutomation.Core.Point(4, 0),
            };
            var poly = page.DrawPolyline(pts);
            // </sample>
        }
    }
}
