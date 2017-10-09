using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioConnection)]
    public class NewVisioConnection : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Shape[] From { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public IVisio.Shape[] To { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = false)]
        public IVisio.Master Master { get; set; }

        protected override void ProcessRecord()
        {
            var connectors = this.Client.Connection.Connect(this.From, this.To, this.Master);
            this.WriteObject(connectors, true);
        }
    }
}