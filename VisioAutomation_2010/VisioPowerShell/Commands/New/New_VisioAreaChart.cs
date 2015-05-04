using VisioAutomation.Drawing;
using VisioAutomation.Models.Charting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioAreaChart")]
    public class New_VisioAreaChart : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [SMA.ParameterAttribute(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [SMA.ParameterAttribute(Position = 2, Mandatory = true)]
        public double X1 { get; set; }

        [SMA.ParameterAttribute(Position = 3, Mandatory = true)]
        public double Y1 { get; set; }

        [SMA.ParameterAttribute(Position = 4, Mandatory = true)]
        public double[] Values;

        [SMA.ParameterAttribute(Mandatory = false)]
        public string[] Labels;

        protected override void ProcessRecord()
        {
            var rect = this.GetRectangle();
            var chart = new AreaChart(rect);
            chart.DataPoints = new DataPointList(this.Values, this.Labels);
            this.WriteObject(chart);
        }

        protected Rectangle GetRectangle()
        {
            return new Rectangle(this.X0, this.Y0, this.X1, this.Y1);
        }
    }
}