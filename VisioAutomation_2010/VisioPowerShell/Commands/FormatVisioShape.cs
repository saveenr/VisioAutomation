using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Format, VisioPowerShell.Commands.Nouns.VisioShape)]
    public class FormatVisioShape : VisioCmdlet
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
        public VisioScripting.Models.AlignmentVertical? AlignVertical = null;

        [SMA.Parameter(Mandatory = false)]
        public VisioScripting.Models.AlignmentHorizontal? AlignHorizontal = null;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targets = new VisioScripting.Models.TargetShapes(this.Shapes);

            if (this.NudgeX != 0.0 || this.NudgeY != 0.0)
            {
                this.Client.Arrange.NudgeShapes(targets, this.NudgeX, this.NudgeY);
            }

            if (this.DistributeHorizontal)
            {
                this.Client.Distribute.DistributeShapesOnAxis(targets, VisioScripting.Models.Axis.XAxis);
            }

            if (this.DistributeVertical)
            {
                this.Client.Distribute.DistributeShapesOnAxis(targets, VisioScripting.Models.Axis.YAxis);
            }

            if (this.AlignVertical.HasValue)
            {
                this.Client.Align.AlignSelectedShapesVertical(targets, this.AlignVertical.Value);
            }

            if (this.AlignHorizontal.HasValue)
            {
                this.Client.Align.AlignSelectedShapesHorizontal(targets, this.AlignHorizontal.Value);
            }

        }
    }
}