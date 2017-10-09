using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioDocument)]
    public class GetVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        [SMA.ValidateNotNullOrEmpty]
        public string Name = null;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ActiveDocument;

        protected override void ProcessRecord()
        {
            var application = this.Client.Application.Get();

            if (this.ActiveDocument)
            {
                var active_doc = application.ActiveDocument;
                this.WriteObject(active_doc);
                return;
            }

            var docs = this.Client.Document.GetDocumentsByName(this.Name);
            this.WriteObject(docs, true);
        }
    }
}