using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioMaster
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioMaster)]
    public class NewVisioMaster : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public string Name;

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document Document;

        protected override void ProcessRecord()
        {
            var master = this.Client.Master.NewMaster(this.Document, this.Name);
            this.WriteObject(master);
        }
    }
}