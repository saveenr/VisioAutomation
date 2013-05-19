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
            var doc = this.ScriptingSession.Document.New();
            this.WriteObject(doc);
        }
    }
}