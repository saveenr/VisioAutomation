using VAS=VisioAutomation.Scripting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Set", "Zoom")]
    public class Set_Zoom : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)] public VA.Scripting.Zoom ZoomLevel =
            VA.Scripting.Zoom.In;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.View.Zoom(this.ZoomLevel);
        }
    }
}