using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;

using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public void Align(VA.Drawing.AlignmentHorizontal flag)
        {
            this.ScriptingSession.Layout.Align(flag);
        }

        public void Align(VA.Drawing.AlignmentHorizontal flag, double x)
        {
            this.ScriptingSession.Layout.Align(flag, x);
        }

        public void Align(VA.Drawing.AlignmentVertical flag)
        {
            this.ScriptingSession.Layout.Align(flag);
        }

        public void Align(VA.Drawing.AlignmentVertical flag, double y)
        {
            this.ScriptingSession.Layout.Align(flag, y);
        }
    }
}