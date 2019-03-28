using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Close, Nouns.VisioDocument)]
    public class CloseVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Document[] Documents;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Force;

        protected override void ProcessRecord()
        {
            var t = new VisioScripting.Models.TargetDocuments(this.Documents);
            t.Resolve(this.Client);
            this.Client.Document.CloseDocuments(t, this.Force);
        }
    }
}