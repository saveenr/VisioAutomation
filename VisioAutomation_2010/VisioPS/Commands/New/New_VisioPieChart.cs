using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioPieChart")]
    public class New_VisioPieChart : VisioPS.VisioPSCmdlet
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
            var scriptingsession = this.ScriptingSession;
            var center = new VA.Drawing.Point(this.X0, this.Y0);

            var chart = new VA.Models.Charting.PieChart(center,this.Radius);
            chart.InnerRadius = this.InnerRadius;

            for (int i = 0; i < this.Values.Length; i++)
            {
                var dp = new VA.Models.Charting.DataPoint();
                dp.Value = this.Values[i];
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