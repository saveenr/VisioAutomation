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

            return new TargetDocument(cmdtarget.ActiveDocument);
        }
    }
}