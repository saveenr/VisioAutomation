using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page pages/helper.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class PagesHelperSamples
    {
        public static void Sample_01(IVisio.Page page)
        {
            // <sample id="pages/helper#01" doc="pages/helper.md" line="11">
            VisioAutomation.Pages.PageHelper.SetSize(page, new VisioAutomation.Core.Size(11, 8.5));
            // </sample>
        }

        public static void Sample_02(IVisio.Page page, double x)
        {
            // <sample id="pages/helper#02" doc="pages/helper.md" line="17">
            var size = VisioAutomation.Pages.PageHelper.GetSize(page);
            System.Console.WriteLine($"{size.Width} x {size.Height}");
            // </sample>
        }

        public static void Sample_03(IVisio.Page page)
        {
            // <sample id="pages/helper#03" doc="pages/helper.md" line="28">
            var padding = new VisioAutomation.Core.Size(0.5, 0.5);
            VisioAutomation.Pages.PageHelper.ResizeToFitContents(page, padding);
            // </sample>
        }

        public static void Sample_04(IVisio.Document doc)
        {
            // <sample id="pages/helper#04" doc="pages/helper.md" line="41">
            var src  = doc.Pages[1];
            var dest = doc.Pages.Add();
            VisioAutomation.Pages.PageHelper.Duplicate(src, dest);
            // </sample>
        }
    }
}
