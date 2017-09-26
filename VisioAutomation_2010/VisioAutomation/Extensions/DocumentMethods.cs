using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class DocumentMethods
    {
        public static void Close(this IVisio.Document doc, bool force_close)
        {
            if (force_close)
            {
                var new_alert_response = Application.AlertResponseCode.No;
                var app = doc.Application;

                using (var alertresponse = new Application.AlertResponseScope(app, new_alert_response))
                {
                    doc.Close();
                }
            }
            else
            {
                doc.Close();
            }
        }

        public static IEnumerable<IVisio.Document> ToEnumerable(this IVisio.Documents documents)
        {
            short count = documents.Count;
            for (int i = 0; i < count; i++)
            {
                yield return documents[i + 1];
            }
        }

        public static IList<IVisio.Document> ToList(this IVisio.Documents documents)
        {
            int count = documents.Count;
            var list = new List<IVisio.Document>(count);
            for (int i = 0; i < count; i++)
            {
                list.Add(documents[i+1]);
            }
            return list;
        }

        public static IVisio.Document OpenStencil(this IVisio.Documents documents, string filename)
        {
            var stencil = VisioAutomation.Documents.DocumentHelper.TryOpenStencil(documents, filename);
            if (stencil == null)
            {
                string msg = string.Format("Could not open stencil \"{0}\"", filename);
                throw new VisioAutomation.Exceptions.VisioOperationException(msg);
            }
            return stencil;
        }
    }
}