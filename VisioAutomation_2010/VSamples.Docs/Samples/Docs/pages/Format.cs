using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page pages/format.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class PagesFormatSamples
    {
        public static void Sample_01(IVisio.Page page)
        {
            // <sample id="pages/format#01" doc="pages/format.md" line="37">
            var fmt = new VisioAutomation.Pages.PageFormatCells();
            fmt.Width  = 11;
            fmt.Height = 8.5;

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(fmt);
            writer.Commit(page.PageSheet, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_02(IVisio.Page page)
        {
            // <sample id="pages/format#02" doc="pages/format.md" line="53">
            var fmt = VisioAutomation.Pages.PageFormatCells.GetCells(
                page.PageSheet,
                VisioAutomation.Core.CellValueType.Result);

            System.Console.WriteLine($"Width = {fmt.Width.Value}");
            // </sample>
        }
    }
}
