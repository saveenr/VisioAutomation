using System.Collections.Generic;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioCustomProperty")]
    public class Get_VisioCustomProperty : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IList<Microsoft.Office.Interop.Visio.Shape> Shapes;
        
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var dic = scriptingsession.CustomProp.Get(this.Shapes);
            this.WriteObject(dic);
        }
    }
}