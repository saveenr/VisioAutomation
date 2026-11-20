using IVisio = Microsoft.Office.Interop.Visio;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page visio-scripting/document.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class VisioScriptingDocumentSamples
    {
        public static void Sample_01(VisioScripting.Client client, IVisio.Document stencil, double a)
        {
            // <sample id="visio-scripting/document#01" doc="visio-scripting/document.md" line="36">
            // Create a blank document and a stencil-backed one
            var blank = client.Document.NewDocument();
            var fromTemplate = client.Document.NewDocumentFromTemplate(@"C:\templates\flowchart.vstx");

            // Open an existing drawing
            var existing = client.Document.OpenDocument(@"C:\drawings\plan.vsdx");

            // Save then close everything
            client.Document.SaveDocument(VisioScripting.TargetDocument.Auto);
            client.Document.CloseAllDocumentsWithoutSaving();
            // </sample>
        }
    }
}
