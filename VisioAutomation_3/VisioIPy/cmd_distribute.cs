using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public void Distribute(VA.Drawing.AlignmentHorizontal alignment)
        {
            this.ScriptingSession.Layout.Distribute(alignment);
        }

        public void Distribute(VA.Drawing.AlignmentVertical alignment)
        {
            this.ScriptingSession.Layout.Distribute(alignment);
        }

        public void DistributeSpace(VA.Drawing.Axis axis)
        {
            this.ScriptingSession.Layout.Distribute(axis);
        }

        public void DistributeSpace(VA.Drawing.Axis axis, double d)
        {
            this.ScriptingSession.Layout.Distribute(axis, d);
        }
    }
}