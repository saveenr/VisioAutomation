using VisioScripting;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Format, Nouns.VisioShape)]
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

        protected override void ProcessRecord()
        {
            if (this.NudgeX != 0.0 || this.NudgeY != 0.0)
            {
                this.Client.Arrange.NudgeSelection(new TargetSelection(), this.NudgeX, this.NudgeY);
            }

            if (this.DistributeHorizontal)
            {
                this.Client.Distribute.DistributeSelectionOnAxis(new TargetSelection(), VisioScripting.Models.Axis.XAxis);
            }

            if (this.DistributeVertical)
            {
                this.Client.Distribute.DistributeSelectionOnAxis(new TargetSelection(), VisioScripting.Models.Axis.YAxis);
            }

            if (this.AlignVertical.HasValue)
            {
                this.Client.Align.AlignSelectionVertical(new TargetSelection(), this.AlignVertical.Value);
            }

            if (this.AlignHorizontal.HasValue)
            {
                this.Client.Align.AlignSelectionHorizontal(new TargetSelection(), this.AlignHorizontal.Value);
            }

        }
    }
}