using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetDocument: TargetObject<IVisio.Document>
    {
        private TargetDocument() :base()
        {
        }

        public TargetDocument(IVisio.Document doc) : base(doc)
        {
        }

        public TargetDocument Resolve(VisioScripting.Client client)
        {
            if (this.Resolved)
            {
                return this;
            }

            var flags = CommandTargetFlags.RequireApplication |
                        CommandTargetFlags.RequireDocument |
                        CommandTargetFlags.RequirePage;

            var command_target = new CommandTarget(client, flags);

            client.Output.WriteVerbose("Resolving to active document (name={0})", command_target.ActiveDocument.Name);

            return new TargetDocument(command_target.ActiveDocument);
        }

        public IVisio.Document Document => this._get_item_safe();

        public static TargetDocument Auto => new TargetDocument();

    }
}