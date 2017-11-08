using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioModelPieChart)]
    public class NewVisioModelPieChart : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = true)]
        public double Radius { get; set; }

        [SMA.Parameter(Position = 3, Mandatory = true)]
        public double[] Values;

        [SMA.Parameter(Mandatory = false)]
        public string[] Labels;

        [SMA.Parameter(Mandatory = false)]
        public double InnerRadius = 0;

        protected override void ProcessRecord()
        {
            var center = new VisioAutomation.Geometry.Point(this.X0, this.Y0);
            var chart = new VisioAutomation.Models.Charting.PieChart(center,this.Radius);
            chart.InnerRadius = this.InnerRadius;
            chart.DataPoints = new VisioAutomation.Models.Charting.DataPointList(this.Values, this.Labels);
            this.WriteObject(chart);
        }
    }
}