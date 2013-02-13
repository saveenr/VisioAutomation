using System.Collections.Generic;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioControl")]
    public class Get_VisioControl : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IList<Microsoft.Office.Interop.Visio.Shape> Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var controls = this.ScriptingSession.Control.Get(this.Shapes);

            this.WriteObject(controls);
        }
    }
}