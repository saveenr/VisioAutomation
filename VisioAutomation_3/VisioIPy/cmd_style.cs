using VAS = VisioAutomation.Scripting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;

namespace VisioIPy
{
    public partial class VisioIPySession
    {
        public void SetStyle( string stylename, string fontname)
        {
            this.ScriptingSession.Text.SetStyleProperties(stylename, fontname);
        }
    }
}