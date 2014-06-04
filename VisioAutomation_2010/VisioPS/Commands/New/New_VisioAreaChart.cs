using System.Collections.Generic;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioAreaChart")]
    public class New_VisioAreaChart : VisioCmdlet
    {
        [System.Management.Automation.Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [System.Management.Automation.Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [System.Management.Automation.Parameter(Position = 2, Mandatory = true)]
        public double X1 { get; set; }

        [System.Management.Automation.Parameter(Position = 3, Mandatory = true)]
        public double Y1 { get; set; }

        [SMA.Parameter(Position = 4, Mandatory = true)]
        public double[] Values;

        [SMA.Parameter(Mandatory = false)]
        public string[] Labels;

        protected override void ProcessRecord()
        {
            var rect = this.GetRectangle();
            var chart = new VA.Models.Charting.AreaChart(rect);
            chart.DataPoints = new VA.Models.Charting.DataPointList(this.Values, this.Labels);
            this.WriteObject(chart);
        }

        protected VisioAutomation.Drawing.Rectangle GetRectangle()
        {
            return new VisioAutomation.Drawing.Rectangle(X0, Y0, X1, Y1);
        }
    }
}