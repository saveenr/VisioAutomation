using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands.Close
{
    [Cmdlet(VerbsCommon.Close, "VisioDocument")]
    public class Close_VisioDocument : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Document[] Documents;

        [Parameter(Mandatory = false)]
        public SwitchParameter Force;

        protected override void ProcessRecord()
        {
            if (this.Documents== null)
            {
                var app = this.client.Application.Get();
                var doc = app.ActiveDocument;
                if (doc != null)
                {
                    VA.Documents.DocumentHelper.Close(doc,this.Force);
                }
            }
            else
            {
                foreach (var doc in this.Documents)
                {
                    this.client.WriteVerbose("Closing doc with ID={0} Name={1}", doc.ID,doc.Name);
                    VA.Documents.DocumentHelper.Close(doc, this.Force);
                }
            }
        }
    }
}