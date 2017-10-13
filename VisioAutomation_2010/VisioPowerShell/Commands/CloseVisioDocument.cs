using System.Linq;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Close, VisioPowerShell.Commands.Nouns.VisioDocument)]
    public class CloseVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document[] Documents;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Force;

        protected override void ProcessRecord()
        {
            if (this.Documents == null)
            {
                this.Client.Document.Close(this.Force);
            }
            else
            {
                this.Client.Document.Close(this.Documents.ToList(), this.Force);
            }
        }
    }
}