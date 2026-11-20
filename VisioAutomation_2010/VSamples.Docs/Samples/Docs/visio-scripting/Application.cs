using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/application.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingApplicationSamples
    {
        public static void Sample_01(double a)
        {
            // <sample id="visio-scripting/application#01" doc="visio-scripting/application.md" line="32">
            var client = new VisioScripting.Client(null);

            // Attach a fresh Visio if we have none
            if (!client.Application.HasApplication)
            {
                client.Application.NewApplication();
            }

            // Make sure the attached Visio is still alive (the user may have closed it)
            if (!client.Application.ValidateApplication())
            {
                client.Application.NewApplication();
            }

            System.Console.WriteLine("Running Visio {0}", client.Application.ApplicationVersion);

            client.Application.MoveWindowToFront();
            client.Application.SetWindowRectangle(new System.Drawing.Rectangle(100, 100, 1200, 800));
            // </sample>
        }
    }
}
