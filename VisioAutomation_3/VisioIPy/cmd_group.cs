using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IVisio.Shape Group()
        {
            return this.ScriptingSession.Layout.Group();
        }

        public void Ungroup()
        {
            this.ScriptingSession.Layout.Ungroup();
        }
    }
}