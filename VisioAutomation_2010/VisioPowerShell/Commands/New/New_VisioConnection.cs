using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioConnection")]
    public class New_VisioConnection : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public IVisio.Shape[] From { get; set; }

        [SMA.ParameterAttribute(Position = 1, Mandatory = true)]
        public IVisio.Shape[] To { get; set; }

        [SMA.ParameterAttribute(Position = 2, Mandatory = false)]
        public IVisio.Master Master { get; set; }

        protected override void ProcessRecord()
        {
            var connectors = this.client.Connection.Connect(this.From, this.To, this.Master);
            this.WriteObject(connectors, false);
        }
    }
}