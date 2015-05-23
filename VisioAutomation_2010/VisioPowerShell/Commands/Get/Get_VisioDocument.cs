using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Get, "VisioDocument")]
    public class Get_VisioDocument : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = false)]
        [SMA.ValidateNotNullOrEmptyAttribute]
        public string Name = null;

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter ActiveDocument;

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