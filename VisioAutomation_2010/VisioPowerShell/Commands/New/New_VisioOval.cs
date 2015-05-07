using SMA = System.Management.Automation;
using VA=VisioAutomation;
namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioOval")]
    public class New_VisioOval : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [SMA.ParameterAttribute(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [SMA.ParameterAttribute(Position = 2, Mandatory = true)]
        public double X1 { get; set; }

        [SMA.ParameterAttribute(Position = 3, Mandatory = true)]
        public double Y1 { get; set; }

        protected override void ProcessRecord()
        {
            var rect = this.GetRectangle();
            var shape = this.client.Draw.Oval(rect.Left, rect.Bottom, rect.Right, rect.Top);
            this.WriteObject(shape);
        }

        protected VA.Drawing.Rectangle GetRectangle()
        {
            return new VA.Drawing.Rectangle(this.X0, this.Y0, this.X1, this.Y1);
        }
    }
}