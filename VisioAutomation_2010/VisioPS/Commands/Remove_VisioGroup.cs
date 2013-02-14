using System.Collections.Generic;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioGroup")]
    public class Remove_VisioGroup : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IList<Microsoft.Office.Interop.Visio.Shape> Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Layout.Ungroup(this.Shapes);
        }
    }
}