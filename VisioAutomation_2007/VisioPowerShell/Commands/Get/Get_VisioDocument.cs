using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioDocument")]
    public class Get_VisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        [SMA.ValidateNotNullOrEmpty]
        public string Name = null;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ActiveDocument;

        protected override void ProcessRecord()
        {
            var application = this.client.VisioApplication;

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