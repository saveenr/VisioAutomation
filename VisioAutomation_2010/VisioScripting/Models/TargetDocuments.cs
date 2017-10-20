using VisioScripting.Commands;

namespace VisioScripting.Models
{
    public class TargetDocuments
    {
        public Microsoft.Office.Interop.Visio.Document[] Documents { get; private set; }

        public TargetDocuments()
        {
            // This explicitly means that the active document will be used
            this.Documents = null;
        }

        public TargetDocuments(params Microsoft.Office.Interop.Visio.Document[] docs)
        {
            // This explicitly means that the active document will be used
            this.Documents = docs;
        }

        public Microsoft.Office.Interop.Visio.Document[] Resolve(VisioScripting.Client client)
        {
            if (this.Documents == null)
            {
                var cmdtarget = client.GetCommandTarget(CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument |
                                                        CommandTargetFlags.ActivePage);
                this.Documents = new[] { cmdtarget.ActiveDocument };
            }

            return this.Documents;
        }
    }
}