using System.Collections.Generic;
using VA= VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioDistributeShape")]
    public class Invoke_VisioDistributeShape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VA.Drawing.Axis Axis { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = false)] 
        public double Distance = -1.0;

        [SMA.Parameter(Mandatory = false)]
       public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.Distance < 0)
            {
                scriptingsession.Layout.Distribute(this.Shapes, this.Axis);
            }
            else
            {
                scriptingsession.Layout.Distribute(this.Shapes, this.Axis, this.Distance);
            }
        }
    }
}