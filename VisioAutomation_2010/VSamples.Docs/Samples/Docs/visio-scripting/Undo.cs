using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/undo.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingUndoSamples
    {
        public static void Sample_01(VisioScripting.Client client, double a)
        {
            // <sample id="visio-scripting/undo#01" doc="visio-scripting/undo.md" line="17">
            // Wrap a custom batch into a single Undo step
            using (var scope = client.Undo.NewUndoScope("Add labelled rectangle"))
            {
                var rect = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 0, 0, 4, 2);
                client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, rect);
                client.Text.SetShapeText(VisioScripting.TargetShapes.Auto, new[] { "Hello" });
            }

            // Now a single Ctrl+Z (or this call) reverses the whole batch
            client.Undo.UndoLastAction();
            client.Undo.RedoLastAction();
            // </sample>
        }
    }
}
