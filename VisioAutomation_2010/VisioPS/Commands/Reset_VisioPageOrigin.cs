using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Reset, "VisioPageOrigin")]
    public class Reset_VisioPageOrigin : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            ScriptingSession.Page.ResetOrigin();
        }
    }
}