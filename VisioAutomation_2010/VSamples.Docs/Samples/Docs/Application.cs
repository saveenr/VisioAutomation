using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page application.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ApplicationSamples
    {
        public static void Sample_01(IVisio.Application app)
        {
            // <sample id="application#01" doc="application.md" line="9">
            var version = VisioAutomation.Application.ApplicationHelper.GetVersion(app);

            if (version.Major >= 15)
            {
                // Visio 2013 or later
            }
            // </sample>
        }

        public static void Sample_02(IVisio.Application app)
        {
            // <sample id="application#02" doc="application.md" line="22">
            VisioAutomation.Application.ApplicationHelper.Quit(app, force_close: true);
            // </sample>
        }

        public static void Sample_03(IVisio.Application app)
        {
            // <sample id="application#03" doc="application.md" line="32">
            string content_dir = VisioAutomation.Application.ApplicationHelper.GetContentLocation(app);
            // e.g. C:\Program Files\Microsoft Office\root\Office16\Visio Content\1033
            // </sample>
        }

        public static void Sample_04(IVisio.Application app)
        {
            // <sample id="application#04" doc="application.md" line="43">
            VisioAutomation.Application.ApplicationHelper.BringWindowToTop(app);
            // </sample>
        }
    }
}
