using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
{
    public class TargetDocuments
    {
        public IVisio.Document[] Documents { get; private set; }

        public TargetDocuments()
        {
            // This explicitly means that the active document will be used
            this.Documents = null;
        }

        public TargetDocuments(params IVisio.Document[] docs)
        {
            // This explicitly means that the active document will be used
            this.Documents = docs;
        }

        public IVisio.Document[] Resolve(VisioScripting.Client client)
        {
            if (this.Documents == null)
            {
                var cmdtarget = client.GetCommandTarget(
                    Commands.CommandTargetFlags.Application | 
                    Commands.CommandTargetFlags.ActiveDocument |
                    Commands.CommandTargetFlags.ActivePage);
                this.Documents = new[] { cmdtarget.ActiveDocument };
            }

            return this.Documents;
        }
    }
}