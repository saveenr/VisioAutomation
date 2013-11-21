using System.Collections.Generic;
using VA= VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioDistributeShape")]
    public class Invoke_VisioDistributeShape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Horizontal { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Vertical { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.Horizontal)
            {
                scriptingsession.Layout.Distribute(this.Shapes, VA.Drawing.Axis.XAxis);                
            }

            if (this.Vertical)
            {
                scriptingsession.Layout.Distribute(this.Shapes, VA.Drawing.Axis.YAxis);
            }

        }
    }
}