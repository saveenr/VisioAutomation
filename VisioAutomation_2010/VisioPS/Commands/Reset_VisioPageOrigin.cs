using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Reset, "VisioPageOrigin")]
    public class Reset_VisioPageOrigin : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page Page;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            ScriptingSession.Page.ResetOrigin(this.Page);
        }
    }
}