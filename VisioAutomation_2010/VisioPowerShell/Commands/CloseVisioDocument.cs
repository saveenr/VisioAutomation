using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Close, VisioPowerShell.Commands.Nouns.VisioDocument)]
    public class CloseVisioDocument : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Document[] Documents;

        [Parameter(Mandatory = false)]
        public SwitchParameter Force;

        protected override void ProcessRecord()
        {
            if (this.Documents== null)
            {
                var app = this.Client.Application.Get();
                var doc = app.ActiveDocument;
                if (doc != null)
                {
                    doc.Close(this.Force);
                }
            }
            else
            {
                foreach (var doc in this.Documents)
                {
                    this.Client.WriteVerbose("Closing doc with ID={0} Name={1}", doc.ID,doc.Name);
                    doc.Close(this.Force);
                }
            }
        }
    }
}