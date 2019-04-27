using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Format, Nouns.VisioShape)]
    public class FormatVisioShape : VisioCmdlet
    {
        // TODO: This cmdlet only affects the active selection but acts like it can handle non-selections due to its names
        // the same is true for the commands it calls

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

        // PSEUDOCONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            this.HandlePseudoContext(this.Shape);

            if (this.NudgeX != 0.0 || this.NudgeY != 0.0)
            {
                this.Client.Arrange.Nudge(VisioScripting.TargetSelection.Auto, this.NudgeX, this.NudgeY);
            }

            if (this.DistributeHorizontal)
            {
                this.Client.Arrange.DistributeOnAxis(VisioScripting.TargetSelection.Auto, VisioScripting.Models.Axis.XAxis);
            }

            if (this.DistributeVertical)
            {
                this.Client.Arrange.DistributeOnAxis(VisioScripting.TargetSelection.Auto, VisioScripting.Models.Axis.YAxis);
            }

            if (this.AlignVertical.HasValue)
            {
                this.Client.Arrange.AlignVertical(VisioScripting.TargetSelection.Auto, this.AlignVertical.Value);
            }

            if (this.AlignHorizontal.HasValue)
            {
                this.Client.Arrange.AlignHorizontal(VisioScripting.TargetSelection.Auto, this.AlignHorizontal.Value);
            }

        }
    }
}