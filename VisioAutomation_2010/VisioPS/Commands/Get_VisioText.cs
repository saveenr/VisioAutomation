using VAS =VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioText")]
    public class Get_VisioText : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var t = scriptingsession.Text.GetText();
            this.WriteObject(t);
        }
    }
}