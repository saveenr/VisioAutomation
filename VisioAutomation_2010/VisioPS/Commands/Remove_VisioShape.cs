using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioShape")]
    public class Remove_VisioShape : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Selection.Delete();
        }
    }
}