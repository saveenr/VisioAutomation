using VisioAutomation.Drawing;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Format, "VisioShape")]
    public class Format_VisioShape : VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = false)]
        public double NudgeX { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public double NudgeY { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter DistributeHorizontal { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter DistributeVertical { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public VerticalAlignment AlignVertical = VerticalAlignment.None;

        [SMA.ParameterAttribute(Mandatory = false)]
        public HorizontalAlignment AlignHorizontal = HorizontalAlignment.None;

        [SMA.ParameterAttribute(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            if (this.NudgeX != 0.0 || this.NudgeY != 0.0)
            {
                this.client.Arrange.Nudge(this.Shapes, this.NudgeX, this.NudgeY);                
            }

            if (this.DistributeHorizontal)
            {
                this.client.Arrange.Distribute(this.Shapes, Axis.XAxis);
            }

            if (this.DistributeVertical)
            {
                this.client.Arrange.Distribute(this.Shapes, Axis.YAxis);
            }

            if (this.AlignVertical != VerticalAlignment.None)
            {
                this.client.Arrange.Align(this.Shapes, (AlignmentVertical) this.AlignVertical);
            }

            if (this.AlignHorizontal != HorizontalAlignment.None)
            {
                this.client.Arrange.Align(this.Shapes, (AlignmentHorizontal) this.AlignHorizontal);
            }

        }
    }
}