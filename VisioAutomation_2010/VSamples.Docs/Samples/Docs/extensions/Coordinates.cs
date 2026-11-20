using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page extensions/coordinates.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ExtensionsCoordinatesSamples
    {
        public static void Sample_01()
        {
            // <sample id="extensions/coordinates#01" doc="extensions/coordinates.md" line="7">

            // </sample>
        }

        public static void Sample_02(IVisio.Page page, IVisio.Shape shape, IVisio.Window window, double a)
        {
            // <sample id="extensions/coordinates#02" doc="extensions/coordinates.md" line="22">
            // Bounding box of every selected shape, on the page
            var rect = window.Selection.GetBoundingBox(IVisio.VisBoundingBoxArgs.visBBoxUprightWH);

            // Convert a click in page coords into the shape's local coords
            var local = shape.XYFromPage(new VisioAutomation.Core.Point(5.0, 4.0));
            // </sample>
        }
    }
}
