using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Update", "VisioShapeSizeFromText")]
    public class Update_VisioShapeSizeFromText : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)] public
            SMA.SwitchParameter FromText;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Text.FitShapeToText();
        }
    }
}