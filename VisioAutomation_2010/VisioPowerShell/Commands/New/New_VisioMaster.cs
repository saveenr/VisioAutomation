using VAS=VisioAutomation.Scripting;
using IVisio=Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioMaster")]
    public class New_VisioMaster : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Document Stencil;

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public string Name;

        protected override void ProcessRecord()
        {
            var master = this.ScriptingSession.Master.New(this.Stencil, this.Name);
            this.WriteObject(master);
        }
    }
}