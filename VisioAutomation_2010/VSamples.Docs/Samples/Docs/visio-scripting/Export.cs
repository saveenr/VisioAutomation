using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/export.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingExportSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Page page)
        {
            // <sample id="visio-scripting/export#01" doc="visio-scripting/export.md" line="17">
            // Active page to PNG
            client.Export.ExportPageToImage(VisioScripting.TargetPage.Auto,
                                              @"C:\out\page.png");

            // Active page to SVG
            client.Export.ExportPageToImage(VisioScripting.TargetPage.Auto,
                                              @"C:\out\page.svg");

            // Just the current selection to embedded-SVG XHTML
            client.Export.ExportSelectionToHtml(VisioScripting.TargetSelection.Auto,
                                                  @"C:\out\selection.html");
            // </sample>
        }
    }
}
