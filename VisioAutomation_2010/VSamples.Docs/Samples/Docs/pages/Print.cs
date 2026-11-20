using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page pages/print.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class PagesPrintSamples
    {
        public static void Sample_01(IVisio.Page page)
        {
            // <sample id="pages/print#01" doc="pages/print.md" line="43">
            var pp = new VisioAutomation.Pages.PagePrintCells();
            pp.LeftMargin   = "0.5 in";
            pp.RightMargin  = "0.5 in";
            pp.TopMargin    = "0.75 in";
            pp.BottomMargin = "0.75 in";

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(pp);
            writer.Commit(page.PageSheet, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_02(IVisio.Page page)
        {
            // <sample id="pages/print#02" doc="pages/print.md" line="57">
            var pp = VisioAutomation.Pages.PagePrintCells.GetCells(
                page.PageSheet,
                VisioAutomation.Core.CellValueType.Result);

            System.Console.WriteLine($"LeftMargin = {pp.LeftMargin.Value}");
            // </sample>
        }
    }
}
