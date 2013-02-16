using System.Collections.Generic;
using VAS=VisioAutomation.Scripting;
using VA=VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioStackShape")]
    public class Invoke_VisioStackShape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VA.Drawing.Axis Axis { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)] public double Distance = 0.0;

        [SMA.Parameter(Mandatory = false)]
        public IList<IVisio.Shape> Shapes;
           
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Layout.Stack(this.Shapes, this.Axis, this.Distance);
        }
    }
}