using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioMaster)]
    public class NewVisioMaster : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public string Name;

        [Parameter(Mandatory = false)]
        public IVisio.Document Document;

        protected override void ProcessRecord()
        {
            var master = this.Client.Master.New(this.Document, this.Name);
            this.WriteObject(master);
        }
    }
}