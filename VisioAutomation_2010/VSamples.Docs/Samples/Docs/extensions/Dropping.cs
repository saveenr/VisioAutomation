using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page extensions/dropping.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ExtensionsDroppingSamples
    {
        public static void Sample_01()
        {
            // <sample id="extensions/dropping#01" doc="extensions/dropping.md" line="7">

            // </sample>
        }

        public static void Sample_02(IVisio.Document doc, IVisio.Page page)
        {
            // <sample id="extensions/dropping#02" doc="extensions/dropping.md" line="20">
            var stencil    = doc.Application.Documents.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];

            var s = page.Drop(rectmaster, new VisioAutomation.Core.Point(2.0, 1.0));
            // </sample>
        }
    }
}
