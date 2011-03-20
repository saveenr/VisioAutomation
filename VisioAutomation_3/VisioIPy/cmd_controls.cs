using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public IList<int> AddControl(VA.Controls.ControlCells control)
        {
            return this.ScriptingSession.Control.AddControl(control);
        }

        public void DeleteControl(int n)
        {
            this.ScriptingSession.Control.DeleteControl(n);
        }

        public IDictionary<IVisio.Shape, IList<VA.Controls.ControlCells>> GetControls()
        {
            var dic = this.ScriptingSession.Control.GetControls();
            return dic;
        }
    }
}