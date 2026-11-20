using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page undo-scope.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class UndoScopeSamples
    {
        public static void Sample_01(IVisio.Application app, System.Collections.Generic.List<IVisio.Shape> shapes)
        {
            // <sample id="undo-scope#01" doc="undo-scope.md" line="9">
            using (var scope = new VisioAutomation.Application.UndoScope(app, "Drop and label shapes"))
            {
                var page = app.ActivePage;
                var s1 = page.DrawRectangle(0, 0, 1, 1);
                s1.Text = "A";
                var s2 = page.DrawRectangle(2, 0, 3, 1);
                s2.Text = "B";
            }
            // After Dispose, the user sees one Undo entry: "Drop and label shapes".
            // </sample>
        }

        public static void Sample_02(IVisio.Application app)
        {
            // <sample id="undo-scope#02" doc="undo-scope.md" line="27">
            using (var scope = new VisioAutomation.Application.UndoScope(app, "Risky edit"))
            {
                try
                {
                    // ... edits ...
                }
                catch
                {
                    scope.Commit = false;     // roll back on exception
                    throw;
                }
            }
            // </sample>
        }
    }
}
