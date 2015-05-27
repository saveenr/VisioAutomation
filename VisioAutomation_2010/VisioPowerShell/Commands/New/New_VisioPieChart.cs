using System.Management.Automation;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, "VisioPieChart")]
    public class New_VisioPieChart : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [Parameter(Position = 2, Mandatory = true)]
        public double Radius { get; set; }

        [Parameter(Position = 3, Mandatory = true)]
        public double[] Values;

        [Parameter(Mandatory = false)]
        public string[] Labels;

        [Parameter(Mandatory = false)]
        public double InnerRadius = 0;

        protected override void ProcessRecord()
        {
            var center = new VisioAutomation.Drawing.Point(this.X0, this.Y0);
            var chart = new VisioAutomation.Models.Charting.PieChart(center,this.Radius);
            chart.InnerRadius = this.InnerRadius;
            chart.DataPoints = new VisioAutomation.Models.Charting.DataPointList(this.Values, this.Labels);
            this.WriteObject(chart);
        }
    }
}