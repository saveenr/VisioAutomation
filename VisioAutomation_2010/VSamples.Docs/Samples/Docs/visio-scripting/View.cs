using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/view.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingViewSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Page page, IVisio.Window window)
        {
            // <sample id="visio-scripting/view#01" doc="visio-scripting/view.md" line="26">
            // Fit the active page in the active window
            client.View.SetZoomToObject(VisioScripting.TargetWindow.Auto,
                                          VisioScripting.Models.ZoomToObject.Page);

            // Then zoom in by 50 percent
            client.View.SetZoomValueRelative(VisioScripting.TargetWindow.Auto, 1.5);

            // And later, set absolute zoom to 100 percent
            client.View.SetZoomValue(VisioScripting.TargetWindow.Auto, 1.0);
            // </sample>
        }
    }
}
