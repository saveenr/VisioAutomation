using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class DocumentMethods
    {
        public static void Close(this Microsoft.Office.Interop.Visio.Document doc, bool force_close)
        {
            Documents.DocumentHelper.Close(doc, force_close);
        }

        public static IEnumerable<Document> ToEnumerable(this Microsoft.Office.Interop.Visio.Documents docs)
        {
            return Documents.DocumentHelper.ToEnumerable(docs);
        }

        public static Microsoft.Office.Interop.Visio.Document OpenStencil(this Microsoft.Office.Interop.Visio.Documents docs, string filename)
        {
            return Documents.DocumentHelper.OpenStencil(docs, filename);
        }

    }
}