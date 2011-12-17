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
            short count = docs.Count;
            for (int i = 0; i < count; i++)
            {
                yield return docs[i + 1];
            }
        }

        public static IVisio.Document OpenStencil(this IVisio.Documents docs, string filename)
        {
            return VA.DocumentHelper.OpenStencil(docs, filename);
        }
    }
}