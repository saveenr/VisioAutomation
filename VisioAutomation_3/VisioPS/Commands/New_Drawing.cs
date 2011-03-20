using VAS=VisioAutomation.Scripting;
using IVisio=Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "Drawing")]
    public class New_Drawing : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var doc = this.ScriptingSession.Document.NewDocument();
            this.WriteObject(doc);
        }
    }
}