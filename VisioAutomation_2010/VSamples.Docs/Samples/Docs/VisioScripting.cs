using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingSamples
    {
        public static void Sample_01()
        {
            // <sample id="visio-scripting#01" doc="visio-scripting.md" line="9">
            var app = new IVisio.Application();
            var client = new VisioScripting.Client(app);
            // </sample>
        }

        public static void Sample_02(VisioScripting.Client client)
        {
            // <sample id="visio-scripting#02" doc="visio-scripting.md" line="18">
            client.Document.NewDocument();
            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto,
                                      new VisioAutomation.Core.Rectangle(0, 0, 4, 2));
            client.Text.SetShapeText(VisioScripting.TargetShapes.Auto,
                                      new[] { "Hello, Visio!" });
            // </sample>
        }

        public static void Sample_03(VisioScripting.Client client, System.Collections.Generic.List<IVisio.Shape> shapes, IVisio.Page somePage)
        {
            // <sample id="visio-scripting#03" doc="visio-scripting.md" line="32">
            // Set text on whatever shapes are selected at call time
            client.Text.SetShapeText(VisioScripting.TargetShapes.Auto,
                                      new[] { "label" });

            // Pass an explicit IVisio.Page when you have one
            client.Page.SetActivePage(somePage);
            // </sample>
        }
    }
}
