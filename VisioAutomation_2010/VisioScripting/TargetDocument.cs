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

            var command_target = new CommandTarget(client, CommandTargetRequirementFlags.RequireApplication |
                                                                    CommandTargetRequirementFlags.RequireActiveDocument |
                                                                    CommandTargetRequirementFlags.RequirePage);

            return new TargetDocument(command_target.ActiveDocument);
        }

        public IVisio.Document Document => this._item;

    }
}