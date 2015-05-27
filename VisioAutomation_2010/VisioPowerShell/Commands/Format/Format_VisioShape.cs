using System.Management.Automation;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Format
{
    [Cmdlet(VerbsCommon.Format, "VisioShape")]
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

        [Parameter(Mandatory = false)]
        public Model.VerticalAlignment AlignVertical = Model.VerticalAlignment.None;

        [Parameter(Mandatory = false)]
        public Model.HorizontalAlignment AlignHorizontal = Model.HorizontalAlignment.None;

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            if (this.NudgeX != 0.0 || this.NudgeY != 0.0)
            {
                this.client.Arrange.Nudge(this.Shapes, this.NudgeX, this.NudgeY);                
            }

            if (this.DistributeHorizontal)
            {
                this.client.Arrange.Distribute(this.Shapes, VA.Drawing.Axis.XAxis);
            }

            if (this.DistributeVertical)
            {
                this.client.Arrange.Distribute(this.Shapes, VA.Drawing.Axis.YAxis);
            }

            if (this.AlignVertical != Model.VerticalAlignment.None)
            {
                this.client.Arrange.Align(this.Shapes, (VA.Drawing.AlignmentVertical)this.AlignVertical);
            }

            if (this.AlignHorizontal != Model.HorizontalAlignment.None)
            {
                this.client.Arrange.Align(this.Shapes, (VA.Drawing.AlignmentHorizontal)this.AlignHorizontal);
            }

        }
    }
}