using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ApplicationMethods
    {
        public static void Quit(this IVisio.Application app, bool force_close)
        {
            Application.ApplicationHelper.Quit(app,force_close);
        }
    }
}