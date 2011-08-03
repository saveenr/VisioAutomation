using VAS=VisioAutomation.Scripting;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("New", "Connection")]
    public class New_Connection : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Shape From { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public IVisio.Shape To { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = true)]
        public IVisio.Master Master { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var connectors = scriptingsession.Connection.Connect(Master, new[] { From }, new[] { To });
            this.WriteObject(connectors);
        }
    }
}