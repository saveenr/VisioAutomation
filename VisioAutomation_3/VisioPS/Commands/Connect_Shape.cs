using VAS=VisioAutomation.Scripting;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Connect", "Shape")]
    public class Connect_Shape : VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Master master;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            scriptingsession.Connection.ConnectShapes(master);                
        }
    }
}