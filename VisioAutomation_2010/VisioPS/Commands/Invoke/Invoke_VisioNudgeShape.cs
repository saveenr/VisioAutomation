using VA = VisioAutomation;
using System.Collections.Generic;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioArrange")]
    public class Invoke_VisioArrange : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public double NudgeX { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public double NudgeY { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter DistributeHorizontal { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter DistributeVertical { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public VerticalAlignment Vertical = VerticalAlignment.None;

        [SMA.Parameter(Mandatory = false)]
        public HorizontalAlignment Horizontal = HorizontalAlignment.None;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            if (this.NudgeX != 0.0 || this.NudgeY != 0.0)
            {
                scriptingsession.Arrange.Nudge(this.Shapes, this.NudgeX, this.NudgeY);                
            }

            if (this.DistributeHorizontal)
            {
                scriptingsession.Arrange.Distribute(this.Shapes, VA.Drawing.Axis.XAxis);
            }

            if (this.DistributeVertical)
            {
                scriptingsession.Arrange.Distribute(this.Shapes, VA.Drawing.Axis.YAxis);
            }

            if (this.Vertical != VerticalAlignment.None)
            {
                scriptingsession.Arrange.Align(this.Shapes, (VA.Drawing.AlignmentVertical)Vertical);
            }

            if (this.Horizontal != HorizontalAlignment.None)
            {
                scriptingsession.Arrange.Align(this.Shapes, (VA.Drawing.AlignmentHorizontal)Horizontal);
            }

        }
    }
}