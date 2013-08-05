using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioCustomProperty")]
    public class Get_VisioCustomProperty : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
       public IVisio.Shape[] Shapes;
        
        protected override void ProcessRecord()
        {
            if (this.Shapes == null)
            {
                this.WriteVerboseEx("Null shapes array passed to Get-VisioCustomProperty");
                return;
            }

            if (this.Shapes.Length == 0)
            {
                this.WriteVerboseEx("0 shapes array passed to Get-VisioCustomProperty");
                return;
            }

            var scriptingsession = this.ScriptingSession;
            var dic = scriptingsession.CustomProp.Get(this.Shapes);
            this.WriteObject(dic);
        }
    }
}