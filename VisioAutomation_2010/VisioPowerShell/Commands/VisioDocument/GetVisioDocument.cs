using VisioPowerShell.Models;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioDocument
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioDocument)]
    public class GetVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter ActiveDocument;

        [SMA.Parameter(Position = 0, Mandatory = false)]
        public string Name = null;


        [SMA.Parameter(Mandatory = false)] public VisioPowerShell.Models.DocumentType? Type;

        protected override void ProcessRecord()
        {
            if (this.ActiveDocument)
            {
                var application = this.Client.Application.GetAttachedApplication();
                var active_doc = application.ActiveDocument;
                this.WriteObject(active_doc);
                return;
            }

            // If the active document is not specified then work on all the pages in the application

            var visdoctype = _get_vis_document_type(this.Type);
            var docs = this.Client.Document.FindDocuments(this.Name, visdoctype);
            this.WriteObject(docs, true);
        }

        private static IVisio.VisDocumentTypes? _get_vis_document_type(DocumentType? doctype)
        {
            if (doctype == null)
            {
                return null;
            }

            if (doctype.Value == DocumentType.Drawing)
            {
                return IVisio.VisDocumentTypes.visTypeDrawing;
            }
            else if (doctype.Value == DocumentType.Stencil)
            {
                return IVisio.VisDocumentTypes.visTypeStencil;
            }
            else if (doctype.Value == DocumentType.Template)
            {
                return IVisio.VisDocumentTypes.visTypeTemplate;
            }

            throw new System.ArgumentOutOfRangeException(nameof(doctype));
        }
    }


}