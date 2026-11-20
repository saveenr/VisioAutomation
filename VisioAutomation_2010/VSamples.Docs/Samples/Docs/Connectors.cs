using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page connectors.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ConnectorsSamples
    {
        public static void Sample_01(IVisio.Document doc, IVisio.Page page)
        {
            // <sample id="connectors#01" doc="connectors.md" line="9">
            var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var dyncon  = stencil.Masters["Dynamic Connector"];

            var shape1 = page.DrawRectangle(0, 0, 1, 1);
            var shape2 = page.DrawRectangle(2, 2, 3, 3);
            var con    = page.Drop(dyncon, 0, 0);

            VisioAutomation.Shapes.ConnectorHelper.ConnectShapes(shape1, shape2, con);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape1, IVisio.Shape shape2)
        {
            // <sample id="connectors#02" doc="connectors.md" line="28">
            VisioAutomation.Shapes.ConnectorHelper.ConnectShapes(shape1, shape2, null, false);
            // </sample>
        }

        public static void Sample_03(IVisio.Document doc, IVisio.Page page, IVisio.Shape shape_a, IVisio.Shape shape_b, IVisio.Shape shape_c, IVisio.Shape shape_d)
        {
            // <sample id="connectors#03" doc="connectors.md" line="38">
            var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var dyncon  = stencil.Masters["Dynamic Connector"];

            var connectors = VisioAutomation.Shapes.ConnectorHelper.ConnectShapes(
                page,
                new[] { shape_a, shape_b, shape_c },
                new[] { shape_b, shape_c, shape_d },
                dyncon);
            // </sample>
        }
    }
}
