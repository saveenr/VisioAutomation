using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page pages/ruler-grid.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class PagesRulerGridSamples
    {
        public static void Sample_01(IVisio.Page page)
        {
            // <sample id="pages/ruler-grid#01" doc="pages/ruler-grid.md" line="33">
            var rg = new VisioAutomation.Pages.PageRulerAndGridCells();
            rg.XGridSpacing = "0.25 in";
            rg.YGridSpacing = "0.25 in";

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(rg);
            writer.Commit(page.PageSheet, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_02(IVisio.Page page)
        {
            // <sample id="pages/ruler-grid#02" doc="pages/ruler-grid.md" line="45">
            var rg = VisioAutomation.Pages.PageRulerAndGridCells.GetCells(
                page.PageSheet,
                VisioAutomation.Core.CellValueType.Result);

            System.Console.WriteLine($"XGridSpacing = {rg.XGridSpacing.Value}");
            // </sample>
        }
    }
}
