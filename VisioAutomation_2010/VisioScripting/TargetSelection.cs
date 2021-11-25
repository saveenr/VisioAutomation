
namespace VisioScripting
{
    public class TargetSelection : TargetObject<IVisio.Selection>
    {

        private TargetSelection() : base()
        {
        }

        private TargetSelection(IVisio.Selection selection) : base(selection)
        {
        }

        public TargetSelection ResolveToSelection(Client client)
        {
            if (this.Resolved)
            {
                return this;
            }

            var cmdtarget = client.GetCommandTarget(CommandTargetFlags.RequirePage);
            var app = cmdtarget.Application;
            var window = app.ActiveWindow;
            var selection = window.Selection;

            client.Output.WriteVerbose("Resolving to selection (numshapes={0}) from active window (caption=\"{1}\")", selection.Count, window.Caption);

            return new TargetSelection(selection);
        }

        public IVisio.Selection Selection => this._get_item_safe();
        public static TargetSelection Auto => new TargetSelection();

    }
}