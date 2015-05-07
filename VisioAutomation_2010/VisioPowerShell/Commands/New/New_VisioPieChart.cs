using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioPieChart")]
    public class New_VisioPieChart : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [SMA.ParameterAttribute(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [SMA.ParameterAttribute(Position = 2, Mandatory = true)]
        public double Radius { get; set; }

        [SMA.ParameterAttribute(Position = 3, Mandatory = true)]
        public double[] Values;

        [SMA.ParameterAttribute(Mandatory = false)]
        public string[] Labels;

        [SMA.ParameterAttribute(Mandatory = false)]
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