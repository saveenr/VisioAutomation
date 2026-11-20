using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/grouping.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingGroupingSamples
    {
        public static void Sample_01(VisioScripting.Client client, System.Collections.Generic.List<IVisio.Shape> shapes, IVisio.Shape shapeA, IVisio.Shape shapeB, IVisio.Shape shapeC)
        {
            // <sample id="visio-scripting/grouping#01" doc="visio-scripting/grouping.md" line="24">
            // Select some shapes, then group them
            client.Selection.SelectShapes(VisioScripting.TargetWindow.Auto,
                                            new[] { shapeA, shapeB, shapeC });
            var group = client.Grouping.Group(VisioScripting.TargetSelection.Auto);

            System.Console.WriteLine("Created group {0} with {1} child shapes",
                group.Name, group.Shapes.Count);

            // And later, ungroup it
            client.Grouping.Ungroup(new VisioScripting.TargetShapes(group));
            // </sample>
        }
    }
}
