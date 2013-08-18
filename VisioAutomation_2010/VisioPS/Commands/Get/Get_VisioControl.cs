using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioControl")]
    public class Get_VisioControl : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var controls = this.ScriptingSession.Control.Get(this.Shapes);
            this.WriteObject(controls);
        }
    }
}