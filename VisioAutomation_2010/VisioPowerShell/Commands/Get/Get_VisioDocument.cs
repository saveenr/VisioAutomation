using System.Management.Automation;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, "VisioDocument")]
    public class Get_VisioDocument : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = false)]
        [ValidateNotNullOrEmpty]
        public string Name = null;

        [Parameter(Mandatory = false)]
        public SwitchParameter ActiveDocument;

        protected override void ProcessRecord()
        {
            var application = this.client.Application.Get();

            if (this.ActiveDocument)
            {
                var active_doc = application.ActiveDocument;
                this.WriteObject(active_doc);
                return;
            }

            var docs = this.client.Document.GetDocumentsByName(this.Name);
            this.WriteObject(docs, true);
        }
    }
}