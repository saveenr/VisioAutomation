using VisioScripting.Commands;

namespace VisioScripting.Models
{
    public class TargetDocument
    {
        public Microsoft.Office.Interop.Visio.Document Document { get; private set; }

        public TargetDocument()
        {
            // This explicitly means that the active document will be used
            this.Document = null;
        }

        public TargetDocument(Microsoft.Office.Interop.Visio.Document doc)
        {
            // This explicitly means that the active document will be used
            this.Document = doc;
        }

        public Microsoft.Office.Interop.Visio.Document Resolve(VisioScripting.Client client)
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