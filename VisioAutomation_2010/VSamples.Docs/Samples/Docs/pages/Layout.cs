using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page pages/layout.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class PagesLayoutSamples
    {
        public static void Sample_01(IVisio.Page page)
        {
            // <sample id="pages/layout#01" doc="pages/layout.md" line="63">
            var pl = new VisioAutomation.Pages.PageLayoutCells();
            pl.AvenueSizeX = "0.5 in";
            pl.AvenueSizeY = "0.25 in";
            pl.EnableGrid  = 0;

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(pl);
            writer.Commit(page.PageSheet, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_02(IVisio.Page page)
        {
            // <sample id="pages/layout#02" doc="pages/layout.md" line="76">
            var pl = VisioAutomation.Pages.PageLayoutCells.GetCells(
                page.PageSheet,
                VisioAutomation.Core.CellValueType.Formula);

            System.Console.WriteLine($"AvenueSizeX = {pl.AvenueSizeX.Value}");
            // </sample>
        }
    }
}
