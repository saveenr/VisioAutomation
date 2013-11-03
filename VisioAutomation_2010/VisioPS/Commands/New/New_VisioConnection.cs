using VAS=VisioAutomation.Scripting;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioConnection")]
    public class New_VisioConnection : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Shape[] From { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public IVisio.Shape[] To { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = false)]
        public IVisio.Master Master { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            var connectors = scriptingsession.Connection.Connect(From , To, Master);
            this.WriteObject(connectors, false);
        }
    }
}