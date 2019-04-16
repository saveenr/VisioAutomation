using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetDocuments : TargetObjects<IVisio.Document>
    {

        public TargetDocuments() : base()
        {
        }

        public TargetDocuments(List<IVisio.Document> docs) : base(docs)
        {
        }

        public TargetDocuments(params IVisio.Document[] docs) : base (docs)
        {
        }

        public TargetDocuments Resolve(VisioScripting.Client client)
        {
            if (this.Resolved)
            {
                return this;
            }

            // Handle the unresolved case

            var flags = CommandTargetRequirementFlags.RequireApplication |
                        CommandTargetRequirementFlags.RequireActiveDocument |
                        CommandTargetRequirementFlags.RequirePage;

            var cmdtarget = new CommandTarget(client, flags);

            var docs = new List<IVisio.Document> { cmdtarget.ActiveDocument };

            client.Output.WriteVerbose("Resolving to active document (name={0})", cmdtarget.ActiveDocument.Name);
            return new TargetDocuments(docs);
        }

        public IList<IVisio.Document> Documents => this._get_items_safe();
    }
}