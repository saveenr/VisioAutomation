using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/selection.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingSelectionSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Page page, System.Collections.Generic.List<IVisio.Shape> shapes, double a)
        {
            // <sample id="visio-scripting/selection#01" doc="visio-scripting/selection.md" line="42">
            // Select everything on the active page, snapshot it, then delete a subset
            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);

            var selected = client.Selection.GetSelectedShapes(VisioScripting.TargetWindow.Auto);
            System.Console.WriteLine("Page has {0} shapes", selected.Count);

            // Narrow the selection back down to two specific shapes
            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapes(VisioScripting.TargetWindow.Auto,
                                            new[] { selected[0], selected[1] });

            // Then duplicate them
            client.Selection.DuplicateShapes(VisioScripting.TargetSelection.Auto);
            // </sample>
        }
    }
}
