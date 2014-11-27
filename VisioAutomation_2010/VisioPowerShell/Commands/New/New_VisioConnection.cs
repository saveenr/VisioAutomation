using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioConnection")]
    public class New_VisioConnection : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Shape[] From { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public IVisio.Shape[] To { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = false)]
        public IVisio.Master Master { get; set; }

        protected override void ProcessRecord()
        {
            var connectors = this.client.Connection.Connect(From, To, Master);
            this.WriteObject(connectors, false);
        }
    }
}