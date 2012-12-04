using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsLifecycle.Invoke, "VisioConnect")]
    public class Invoke_VisioConnect : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Master Master;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            scriptingsession.Connection.Connect(Master);                
        }
    }
}