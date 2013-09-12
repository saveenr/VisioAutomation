using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioBarChart")]
    public class New_VisioAreaChart : VisioPS.VisioPSCmdlet
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
            var scriptingsession = this.ScriptingSession;
            var rect = new VA.Drawing.Rectangle(X0, Y0, X1, Y1);
            var center = new VA.Drawing.Point(this.X0, this.Y0);

            var chart = new VA.Models.Charting.AreaChart(rect);

            for (int i = 0; i < this.Values.Length; i++)
            {
                var dp = new VA.Models.Charting.DataPoint(this.Values[i]);
                if (i < this.Labels.Length)
                {
                    dp.Label = this.Labels[i];
                }

                chart.DataPoints.Add(dp);
            }

            this.WriteObject(chart);
        }
    }
}