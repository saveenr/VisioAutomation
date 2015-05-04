using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioLine")]
    public class New_VisioLine : VisioCmdlet
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
            var line = this.client.Draw.Line(this.X0, this.Y0, this.X1, this.Y1);
            this.WriteObject(line);
        }
    }
}