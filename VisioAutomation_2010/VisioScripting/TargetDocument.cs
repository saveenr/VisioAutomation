using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetDocument: TargetObject<IVisio.Document>
    {
        public TargetDocument() :base()
        {
        }

        public TargetDocument(IVisio.Document doc) : base(doc)
        {
        }

        public TargetDocument(IVisio.Document doc, bool isresolved) : base(doc, isresolved)
        {
        }

        public TargetDocument Resolve(VisioScripting.Client client)
        {
            if (this.IsResolved)
            {
                return this;
            }
            var cmdtarget = client.GetCommandTarget(
                Commands.CommandTargetRequirementFlags.RequireApplication |
                Commands.CommandTargetRequirementFlags.RequireActiveDocument |
                Commands.CommandTargetRequirementFlags.RequirePage);

            // It doesn't matter if there is an active document or not
            // at this point it is considered resolved
            return new TargetDocument(cmdtarget.ActiveDocument, true);
        }
    }
}