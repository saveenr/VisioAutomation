using System.Collections.Generic;
using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioConnectionPoint")]
    public class Get_VisioConnectionPoint : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] public IList<IVisio.Shape> Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var dic = scriptingsession.ConnectionPoint.Get(this.Shapes);
            this.WriteObject(dic);
        }
    }
}