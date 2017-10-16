using VisioPowerShell.Models;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Models
{
    public enum DocumentType
    {
        Drawing,
        Stencil,
        Template
    }
}

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioDocument)]
    public class GetVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Name = null;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ActiveDocument;

        [SMA.Parameter(Mandatory = false)] public VisioPowerShell.Models.DocumentType? Type;

        protected override void ProcessRecord()
        {
            if (this.ActiveDocument)
            {
                var application = this.Client.Application.GetActiveApplication();
                var active_doc = application.ActiveDocument;
                this.WriteObject(active_doc);
                return;
            }

            IVisio.VisDocumentTypes? visdoctype = null;

            if (this.Type == DocumentType.Drawing)
            {
                visdoctype = IVisio.VisDocumentTypes.visTypeDrawing;
            }
            else if (this.Type == DocumentType.Stencil)
            {
                visdoctype = IVisio.VisDocumentTypes.visTypeStencil;
            }
            else if (this.Type == DocumentType.Template)
            {
                visdoctype = IVisio.VisDocumentTypes.visTypeTemplate;
            }

            var docs = this.Client.Document.FindDocuments(this.Name, visdoctype);
            this.WriteObject(docs, true);
        }
    }


}