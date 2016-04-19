using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class DocumentMethods
    {
        public static void Close(this IVisio.Document doc, bool force_close)
        {
            Documents.DocumentHelper.Close(doc, force_close);
        }

        public static IEnumerable<IVisio.Document> ToEnumerable(this IVisio.Documents docs)
        {
            return Documents.DocumentHelper.ToEnumerable(docs);
        }

        public static IVisio.Document OpenStencil(this IVisio.Documents docs, string filename)
        {
            return Documents.DocumentHelper.OpenStencil(docs, filename);
        }

    }
}