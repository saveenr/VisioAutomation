using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page stencils-and-masters.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class StencilsAndMastersSamples
    {
        public static void Sample_01()
        {
            // <sample id="stencils-and-masters#01" doc="stencils-and-masters.md" line="5">

            // </sample>
        }

        public static void Sample_02(IVisio.Document doc)
        {
            // <sample id="stencils-and-masters#02" doc="stencils-and-masters.md" line="13">
            var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];
            var page = doc.Pages.Add();

            var shape1 = page.Drop(rectmaster, 1.0, 2.0);

            var p = new VisioAutomation.Core.Point(5.0, 4.0);
            var shape2 = page.Drop(rectmaster, p);
            // </sample>
        }

        public static void Sample_03(IVisio.Document doc)
        {
            // <sample id="stencils-and-masters#03" doc="stencils-and-masters.md" line="26">
            var stencil = doc.Application.Documents.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];
            var page = doc.Pages.Add();

            var centerpoints = new[] {
                                           new VisioAutomation.Core.Point(1, 2),
                                           new VisioAutomation.Core.Point(5, 4)
                                       };
            var masters = new[] { rectmaster, rectmaster };
            short[] shapeids = page.DropManyU(masters, centerpoints);
            // </sample>
        }
    }
}
