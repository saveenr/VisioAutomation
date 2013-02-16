using System.Collections.Generic;
using VAS =VisioAutomation.Scripting;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioText")]
    public class Get_VisioText : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IList<IVisio.Shape> Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var t = scriptingsession.Text.GetText(this.Shapes);
            this.WriteObject(t);
        }
    }
}