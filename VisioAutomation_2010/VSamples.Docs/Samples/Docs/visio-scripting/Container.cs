using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/container.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingContainerSamples
    {
        public static void Sample_01(VisioScripting.Client client, System.Collections.Generic.List<IVisio.Shape> shapes, double a)
        {
            // <sample id="visio-scripting/container#01" doc="visio-scripting/container.md" line="16">
            // Select some shapes first via client.Selection, then drop a container around them
            var selectionWindow = VisioScripting.TargetWindow.Auto;
            client.Selection.SelectAllShapes(selectionWindow);

            // Use a built-in container by name
            var container = client.Container.DropContainer(VisioScripting.TargetPage.Auto,
                                                            "Plain");

            System.Console.WriteLine("Container {0} now wraps the selection", container.Name);
            // </sample>
        }
    }
}
