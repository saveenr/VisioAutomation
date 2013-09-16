using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioGroup")]
    public class New_VisioGroup : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var group = scriptingsession.Layout.Group();
            this.WriteObject(group);
        }
    }
}