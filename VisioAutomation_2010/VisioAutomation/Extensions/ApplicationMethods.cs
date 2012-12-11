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

        public static VA.Application.UndoScope CreateUndoScope(this IVisio.Application app, string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new System.ArgumentOutOfRangeException(name, "name");
            }
            return new VA.Application.UndoScope(app, name);
        }
        
        public static VA.Application.AlertResponseScope CreateAlertResponseScope(this IVisio.Application app, VA.Application.AlertResponseCode code)
        {
            return new VA.Application.AlertResponseScope(app, code);
        }
    }
}