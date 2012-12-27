using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static class ApplicationMethods
    {
        public static void Quit(this IVisio.Application app, bool force_close)
        {
            VA.Application.ApplicationHelper.Quit(app,force_close);
        }
    }
}