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

            var cmdtarget = client.GetCommandTarget(
                Commands.CommandTargetRequirementFlags.RequireApplication | 
                Commands.CommandTargetRequirementFlags.RequireActiveDocument |
                Commands.CommandTargetRequirementFlags.RequirePage);

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
    }
}