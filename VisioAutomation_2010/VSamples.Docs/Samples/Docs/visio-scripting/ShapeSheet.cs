using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/shape-sheet.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingShapeSheetSamples
    {
        public static void Sample_01(VisioScripting.Client client, double a)
        {
            // <sample id="visio-scripting/shape-sheet#01" doc="visio-scripting/shape-sheet.md" line="17">
            // Use the writer to set Width and Height on every selected shape in one commit
            var writer = client.ShapeSheet.GetWriterForPage(VisioScripting.TargetPage.Auto);
            foreach (var shape in client.Selection.GetSelectedShapes(VisioScripting.TargetWindow.Auto))
            {
                writer.SetFormula((short)shape.ID, VisioAutomation.Core.SrcConstants.XFormWidth, "2.0 in");
                writer.SetFormula((short)shape.ID, VisioAutomation.Core.SrcConstants.XFormHeight, "1.0 in");
            }
            writer.Commit();

            // Use the reader to fetch the resolved values back
            var reader = client.ShapeSheet.GetReaderForPage(VisioScripting.TargetPage.Auto);
            // ... (set up queries on `reader`, call reader.Read())

            // Rename the same selection
            client.ShapeSheet.SetShapeName(VisioScripting.TargetShapes.Auto,
                                             new[] { "node-a", "node-b", "node-c" });
            // </sample>
        }
    }
}
