using System.Collections.Generic;
using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioGroup")]
    public class New_VisioGroup : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IList<Microsoft.Office.Interop.Visio.Shape> Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Layout.Group(this.Shapes);
        }
    }
}