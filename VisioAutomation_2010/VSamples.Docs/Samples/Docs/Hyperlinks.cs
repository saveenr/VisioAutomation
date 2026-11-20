using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page hyperlinks.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class HyperlinksSamples
    {
        public static void Sample_01(IVisio.Shape shape)
        {
            // <sample id="hyperlinks#01" doc="hyperlinks.md" line="22">
            var hl = new VisioAutomation.Shapes.HyperlinkCells();
            hl.Address = "https://learn.microsoft.com";
            hl.Description = "Microsoft Learn";

            int row = VisioAutomation.Shapes.HyperlinkHelper.Add(shape, hl);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape, int row)
        {
            // <sample id="hyperlinks#02" doc="hyperlinks.md" line="36">
            var hl = new VisioAutomation.Shapes.HyperlinkCells();
            hl.Address = "https://example.com";
            hl.Description = "Updated description";

            VisioAutomation.Shapes.HyperlinkHelper.Set(shape, (short)row, hl);
            // </sample>
        }

        public static void Sample_03(IVisio.Shape shape)
        {
            // <sample id="hyperlinks#03" doc="hyperlinks.md" line="46">
            int count = VisioAutomation.Shapes.HyperlinkHelper.GetCount(shape);
            // </sample>
        }

        public static void Sample_04(IVisio.Shape shape)
        {
            // <sample id="hyperlinks#04" doc="hyperlinks.md" line="52">
            VisioAutomation.Shapes.HyperlinkHelper.Delete(shape, 0);   // first hyperlink
            // </sample>
        }
    }
}
