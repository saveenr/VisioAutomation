using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page exceptions.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ExceptionsSamples
    {
        public static void Sample_01()
        {
            // <sample id="exceptions#01" doc="exceptions.md" line="9">
            try
            {
                // ... VisioAutomation calls ...
            }
            catch (VisioAutomation.Exceptions.AutomationException ex)
            {
                System.Console.WriteLine($"Library error: {ex.Message}");
            }
            // </sample>
        }

        public static void Sample_02(IVisio.Document doc)
        {
            // <sample id="exceptions#02" doc="exceptions.md" line="24">
            try
            {
                var stencil = doc.Application.Documents.OpenStencil("missing.vss");
            }
            catch (VisioAutomation.Exceptions.VisioOperationException)
            {
                // The stencil could not be opened.
            }
            // </sample>
        }
    }
}
