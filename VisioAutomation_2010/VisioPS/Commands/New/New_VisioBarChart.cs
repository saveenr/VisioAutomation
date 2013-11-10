using VisioAutomation.Models.Charting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioBarChart")]
    public class New_VisioBarChart : RectangleCmdlet
    {
        [SMA.Parameter(Mandatory = true)]
        public double[] Values;

        [SMA.Parameter(Mandatory = false)]
        public string[] Labels;

        protected override void ProcessRecord()
        {
            var rect = this.GetRectangle();
            var chart = new VA.Models.Charting.BarChart(rect);
            chart.DataPoints = new DataPointList(this.Values, this.Labels);
            this.WriteObject(chart);
        }
    }
}