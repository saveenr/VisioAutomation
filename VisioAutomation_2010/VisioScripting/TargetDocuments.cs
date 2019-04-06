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
            if (this.IsResolved)
            {
                return this;
            }

            // Handle the unresolved case

            var cmdtarget = new CommandTarget(client, CommandTargetRequirementFlags.RequireApplication |
                                                                    CommandTargetRequirementFlags.RequireActiveDocument |
                                                                    CommandTargetRequirementFlags.RequirePage);

            if (cmdtarget.ActiveDocument == null)
            {
                var docs = new List<IVisio.Document>(0);
                return new TargetDocuments(docs);
            }
            else
            {
                var docs = new List<IVisio.Document> { cmdtarget.ActiveDocument };
                return new TargetDocuments(docs);
            }
        }

        public IList<IVisio.Document> Documents => this._items;
    }
}