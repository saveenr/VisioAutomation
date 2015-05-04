using VisioAutomation.Documents;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Close, "VisioDocument")]
    public class Close_VisioDocument : VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = false)]
        public IVisio.Document[] Documents;

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter Force;

        protected override void ProcessRecord()
        {
            if (this.Documents== null)
            {
                var app = this.client.VisioApplication;
                var doc = app.ActiveDocument;
                if (doc != null)
                {
                    DocumentHelper.Close(doc,this.Force);
                }
            }
            else
            {
                foreach (var doc in this.Documents)
                {
                    this.client.WriteVerbose("Closing doc with ID={0} Name={1}", doc.ID,doc.Name);
                    DocumentHelper.Close(doc, this.Force);
                }
            }
        }
    }
}