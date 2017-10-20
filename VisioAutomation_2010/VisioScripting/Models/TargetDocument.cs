using VisioScripting.Commands;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
{
    public class TargetDocument
    {
        public IVisio.Document Document { get; private set; }

        public TargetDocument()
        {
            // This explicitly means that the active document will be used
            this.Document = null;
        }

        public TargetDocument(IVisio.Document doc)
        {
            // This explicitly means that the active document will be used
            this.Document = doc;
        }

        public IVisio.Document Resolve(VisioScripting.Client client)
        {
            if (this.Document == null)
            {
                var cmdtarget = client.GetCommandTarget(CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument |
                                                        CommandTargetFlags.ActivePage);
                this.Document = cmdtarget.ActiveDocument;
            }

            return this.Document;
        }
    }
}