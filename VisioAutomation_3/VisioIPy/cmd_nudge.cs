using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public void Nudge(double x, double y)
        {
            this.ScriptingSession.Layout.Nudge(x, y);
        }
    }
}