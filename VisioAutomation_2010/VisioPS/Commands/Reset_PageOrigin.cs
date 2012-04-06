using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Reset, "PageOrigin")]
    public class Reset_PageOrigin : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            ScriptingSession.Page.ResetOrigin();
        }
    }
}