using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page text/character.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class TextCharacterSamples
    {
        public static void Sample_01(IVisio.Shape shape)
        {
            // <sample id="text/character#01" doc="text/character.md" line="52">
            var ch = new VisioAutomation.Text.CharacterCells();
            ch.Font  = "Segoe UI";
            ch.Size  = "14pt";
            ch.Color = "rgb(0,0,128)";
            ch.Style = 1;                 // bold

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(ch);
            writer.Commit(shape, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_02(IVisio.Shape shape)
        {
            // <sample id="text/character#02" doc="text/character.md" line="68">
            var run0 = new VisioAutomation.Text.CharacterCells();
            run0.Font  = "Segoe UI";
            run0.Size  = "10pt";

            var run1 = new VisioAutomation.Text.CharacterCells();
            run1.Font  = "Segoe UI";
            run1.Size  = "10pt";
            run1.Style = 1;                // bold for the second run

            var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
            writer.SetValues(run0, 0);
            writer.SetValues(run1, 1);
            writer.Commit(shape, VisioAutomation.Core.CellValueType.Formula);
            // </sample>
        }

        public static void Sample_03(IVisio.Shape shape)
        {
            // <sample id="text/character#03" doc="text/character.md" line="90">
            // All character rows on one shape:
            var rows = VisioAutomation.Text.CharacterCells.GetCells(
                shape,
                VisioAutomation.Core.CellValueType.Result);

            foreach (var row in rows)
            {
                System.Console.WriteLine($"Row: font={row.Font.Value} size={row.Size.Value}");
            }
            // </sample>
        }
    }
}
