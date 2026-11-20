using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page page-cells.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class PageCellsSamples
    {
        public static void Sample_01(IVisio.Page page, VisioAutomation.Pages.PageFormatCells fmt, VisioAutomation.Pages.PagePrintCells pp, VisioAutomation.Pages.PageRulerAndGridCells rg)
        {
            // <sample id="page-cells#01" doc="page-cells.md" line="26">
            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(fmt);
            writer.SetValues(pp);
            writer.SetValues(rg);
            writer.Commit(page.PageSheet, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }
    }
}
