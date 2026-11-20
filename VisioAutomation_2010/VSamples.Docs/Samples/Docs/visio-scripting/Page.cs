using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/page.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingPageSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Page page, double x, double a)
        {
            // <sample id="visio-scripting/page#01" doc="visio-scripting/page.md" line="42">
            // Create a Letter-sized page in the active document
            var newpage = client.Page.NewPage(VisioScripting.TargetDocument.Auto,
                                              new VisioAutomation.Core.Size(8.5, 11),
                                              isbackgroundpage: false);

            // Switch to landscape orientation across every page in the active document
            client.Page.SetPageOrientation(VisioScripting.TargetPages.Auto,
                                            VisioScripting.Models.PageOrientation.Landscape);

            // Read each page's full dimension record
            var dims = client.Page.GetPageDimensions(VisioScripting.TargetPages.Auto);
            foreach (var d in dims)
            {
                System.Console.WriteLine("Page {0}: {1} x {2} (margins L={3} R={4})",
                    d.PageID, d.PageWidth, d.PageHeight,
                    d.PrintLeftMargin, d.PrintRightMargin);
            }

            // Tighten the active page around its content with a 0.25" border
            client.Page.ResizePageToFitContents(VisioScripting.TargetPages.Auto,
                                                new VisioAutomation.Core.Size(0.25, 0.25));
            // </sample>
        }
    }
}
