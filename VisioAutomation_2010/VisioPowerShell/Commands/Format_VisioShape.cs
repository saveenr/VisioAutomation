using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Format, VisioPowerShell.Commands.Nouns.VisioShape)]
    public class Format_VisioShape : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public double NudgeX { get; set; }

        [Parameter(Mandatory = false)]
        public double NudgeY { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter DistributeHorizontal { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter DistributeVertical { get; set; }

        [Parameter(Mandatory = false)] public double Distance = -1.0;

        [Parameter(Mandatory = false)]
        public VisioScripting.Models.AlignmentVertical? AlignVertical = null;

        [Parameter(Mandatory = false)]
        public VisioScripting.Models.AlignmentHorizontal? AlignHorizontal = null;

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);

            if (this.NudgeX != 0.0 || this.NudgeY != 0.0)
            {
                this.Client.Arrange.Nudge(targets, this.NudgeX, this.NudgeY);
            }

            if (this.DistributeHorizontal)
            {
                if (this.Distance < 0)
                {
                    this.Client.Distribute.DistributeOnAxis(targets, VisioScripting.Models.Axis.XAxis);
                }
                else
                {
                    this.Client.Distribute.DistributeOnAxis(targets, VisioScripting.Models.Axis.XAxis, this.Distance);
                }
            }

            if (this.DistributeVertical)
            {
                if (this.Distance < 0)
                {
                    this.Client.Distribute.DistributeOnAxis(targets, VisioScripting.Models.Axis.YAxis);
                }
                else
                {
                    this.Client.Distribute.DistributeOnAxis(targets, VisioScripting.Models.Axis.YAxis, this.Distance);
                }
            }

            if (this.AlignVertical.HasValue)
            {
                this.Client.Align.AlignVertical(targets, this.AlignVertical.Value);
            }

            if (this.AlignHorizontal.HasValue)
            {
                this.Client.Align.AlignHorizontal(targets, this.AlignHorizontal.Value);
            }

        }
    }
}