using System;
using System.Linq;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioPolyLine")]
    public class New_VisioPolyLine : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double[] Doubles { get; set; }

        protected override void ProcessRecord()
        {
            if ((this.Doubles.Length%2) != 0)
            {
                var exc = new ArgumentOutOfRangeException("polyline has odd number of elements");
                var er = new SMA.ErrorRecord(exc, "POLYLINE_COUNT", SMA.ErrorCategory.InvalidData, null);
                this.WriteError(er);

                return;
            }

            var scriptingsession = this.ScriptingSession;
            var points = VA.Drawing.Point.FromDoubles(this.Doubles).ToList();
            var shape = scriptingsession.Draw.PolyLine(points);
            this.WriteObject(shape);
        }
    }
}