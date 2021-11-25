using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Documents
{
    public static class DocumentHelper
    {

        internal static IVisio.Document TryOpenStencil(IVisio.Documents docs, string filename)
        {
            const short flags = (short)IVisio.VisOpenSaveArgs.visOpenRO | (short)IVisio.VisOpenSaveArgs.visOpenDocked;
            try
            {
                var doc = docs.OpenEx(filename, flags);
                return doc;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
        }

        internal static void Close(IVisio.Document doc, bool force_close)
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

        internal static IVisio.Document OpenStencil(this IVisio.Documents documents, string filename)
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