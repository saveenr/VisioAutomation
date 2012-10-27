using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static class DocumentMethods
    {
        public static void Activate(this IVisio.Document doc)
        {
            VA.Documents.DocumentHelper.Activate(doc);
        }

        public static void Close(this IVisio.Document doc, bool force_close)
        {
            VA.Documents.DocumentHelper.Close(doc, force_close);
        }
    }
}