using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public void Send(VAS.ShapeSendDirection direction)
        {
            this.ScriptingSession.Layout.Send(direction);
        }
    }
}