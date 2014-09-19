using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioSession")]
    public class Get_VisioSession : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var ss = this.ScriptingSession;
            this.WriteObject(ss);
        }
    }
}