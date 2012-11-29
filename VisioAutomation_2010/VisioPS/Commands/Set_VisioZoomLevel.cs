using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioZoomLevel")]
    public class Set_VisioZoomLevel : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double Percent { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            scriptingsession.View.ZoomToPercentage(Percent);
        }
    }
}