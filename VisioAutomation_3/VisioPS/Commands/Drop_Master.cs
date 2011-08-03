using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;
using VA=VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Drop", "Master")]
    public class Drop_Master : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Master[] Masters { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double [] Points { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var points = VA.Drawing.DrawingUtil.DoublesToPoints(Points).ToList();
            var r = scriptingsession.Master.Drop(Masters, points);
            this.WriteObject(r);
        }
    }
}