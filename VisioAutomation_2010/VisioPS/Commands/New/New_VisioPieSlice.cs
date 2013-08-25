using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioPieSlice")]
    public class New_VisioPieSlice : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = true)]
        public double Radius { get; set; }

        [SMA.Parameter(Position = 3, Mandatory = true)]
        public double StartAngle { get; set; }

        [SMA.Parameter(Position = 4, Mandatory = true)]
        public double EndAngle { get; set; }

        [SMA.Parameter(Mandatory = false)] public double InnerRadius = 0;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var center = new VA.Drawing.Point(this.X0, this.Y0);
            if (this.InnerRadius <= 0)
            {
                var shape = scriptingsession.Draw.PieSlice(center, this.Radius, this.StartAngle, this.EndAngle);
                this.WriteObject(shape);
            }
            else
            {
                var shape = scriptingsession.Draw.DoughnutSlice(center, this.InnerRadius, this.Radius, this.StartAngle, this.EndAngle);
                this.WriteObject(shape);                
            }
        }
    }
}