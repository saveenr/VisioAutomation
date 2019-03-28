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
    }
}