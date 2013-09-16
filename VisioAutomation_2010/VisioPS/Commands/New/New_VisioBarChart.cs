using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioBarChart")]
    public class New_VisioBarChart : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = true)]
        public double X1 { get; set; }

        [SMA.Parameter(Position = 3, Mandatory = true)]
        public double Y1 { get; set; }

        [SMA.Parameter(Position = 4, Mandatory = true)]
        public double[] Values;

        [SMA.Parameter(Mandatory = false)]
        public string[] Labels;

        protected override void ProcessRecord()
        {
            var rect = new VA.Drawing.Rectangle(X0, Y0, X1, Y1);
            var chart = new VA.Models.Charting.BarChart(rect);
            chart.DataPoints = VA.Models.Charting.DataPoint.DoublesToDataPoints(this.Values, this.Labels);
            this.WriteObject(chart);
        }
    }
}