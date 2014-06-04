using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioCustomProperty")]
    public class Remove_VisioCustomProperty : VisioPS.VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Name { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.CustomProp.Delete(this.Shapes,Name);
        }
    }
}

