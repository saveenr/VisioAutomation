using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Open, "VisioStencil")]
    public class Open_VisioStencil : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Stencil { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var doc = scriptingsession.Document.OpenStencil(Stencil);
            this.WriteObject(doc);
        }
    }
}