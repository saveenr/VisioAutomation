using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VisioAutomation.Extensions;

namespace VSamples.Docs.Samples.Docs
{
    // Code blocks from the .NET docs page extensions/linq.md. Compile-only: each method is one block,
    // with the names the block assumes (page, shape, client, ...) as parameters.
    public static class ExtensionsLinqSamples
    {
        public static void Sample_01()
        {
            // <sample id="extensions/linq#01" doc="extensions/linq.md" line="7">

            // </sample>
        }

        public static void Sample_02(IVisio.Document doc, IVisio.Page page, double a)
        {
            // <sample id="extensions/linq#02" doc="extensions/linq.md" line="30">
            // Build a font-name to font-ID dictionary
            var fontname_to_id = doc.Fonts.ToEnumerable()
                .ToDictionary(f => f.Name, f => f.ID);

            // Find every page whose name starts with "Backup"
            var backup_pages = doc.Pages.ToEnumerable()
                .Where(p => p.NameU.StartsWith("Backup"))
                .ToList();
            // </sample>
        }
    }
}
