using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VA=VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static class DocumentsMethods
    {
        public static IEnumerable<IVisio.Document> AsEnumerable(this IVisio.Documents docs)
        {
            return docs.Cast<IVisio.Document>();
        }

        public static IVisio.Document OpenStencil(this IVisio.Documents docs, string filename)
        {
            return VA.DocumentHelper.OpenStencil(docs, filename);
        }
    }
}