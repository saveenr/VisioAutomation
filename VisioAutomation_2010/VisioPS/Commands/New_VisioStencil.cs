using VAS=VisioAutomation.Scripting;
using IVisio=Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioStencil")]
    public class New_VisioStencil : VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var doc = this.ScriptingSession.Document.NewStencil();
            this.WriteObject(doc);
        }
    }
}