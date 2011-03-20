using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Extensions
{
    public static class ApplicationMethods
    {
        public static void Quit(this IVisio.Application app, bool force_close)
        {
            VA.ApplicationHelper.Quit(app,force_close);
        }       

        public static UndoScope CreateUndoScope(this IVisio.Application app)
        {
            return new UndoScope(app, "Untitled", VA.UndoCommitFlag.AcceptChanges);
        }

        public static VA.UI.AlertResponseScope CreateAlertResponseScope(this IVisio.Application app, VA.UI.AlertResponseCode code)
        {
            return new VA.UI.AlertResponseScope(app, code);
        }
    }
}