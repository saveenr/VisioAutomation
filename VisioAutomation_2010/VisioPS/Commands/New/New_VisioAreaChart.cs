using System.Collections.Generic;
using VisioAutomation.Models.Charting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioAreaChart")]
    public class New_VisioAreaChart : RectangleCmdlet
    {
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

    }
}