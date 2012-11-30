using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioGroup")]
    public class Remove_VisioGroup : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Layout.Ungroup();
        }
    }
}