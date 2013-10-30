using VAS=VisioAutomation.Scripting;
using IVisio=Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioDocument")]
    public class New_VisioDocument : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            if (!this.ScriptingSession.HasApplication)
            {
                this.ScriptingSession.Application.New();
            }
            else
            {
                if (!this.ScriptingSession.Application.Validate())
                {
                    this.ScriptingSession.Application.New();
                }
            }
            var doc = this.ScriptingSession.Document.New();
            this.WriteObject(doc);
        }
    }
}