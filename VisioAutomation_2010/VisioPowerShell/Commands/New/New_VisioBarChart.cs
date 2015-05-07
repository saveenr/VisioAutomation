using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioBarChart")]
    public class New_VisioBarChart : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [SMA.ParameterAttribute(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [SMA.ParameterAttribute(Position = 2, Mandatory = true)]
        public double X1 { get; set; }

        [SMA.ParameterAttribute(Position = 3, Mandatory = true)]
        public double Y1 { get; set; }

        [SMA.ParameterAttribute(Mandatory = true)]
        public double[] Values;

        [SMA.ParameterAttribute(Mandatory = false)]
        public string[] Labels;

        protected override void ProcessRecord()
        {
            var rect = this.GetRectangle();
            var chart = new VA.Models.Charting.BarChart(rect);
            chart.DataPoints = new VA.Models.Charting.DataPointList(this.Values, this.Labels);
            this.WriteObject(chart);
        }

        protected VA.Drawing.Rectangle GetRectangle()
        {
            return new VA.Drawing.Rectangle(this.X0, this.Y0, this.X1, this.Y1);
        }
    }
}