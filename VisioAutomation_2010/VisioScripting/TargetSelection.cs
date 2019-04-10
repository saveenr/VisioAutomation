using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetSelection : TargetObject<IVisio.Selection>
    {

        public TargetSelection() : base()
        {
        }

        public TargetSelection(IVisio.Selection selection) : base(selection)
        {
        }

        public TargetSelection Resolve(Client client)
        {
            if (this.Resolved)
            {
                return this;
            }

            var cmdtarget = client.GetCommandTargetPage();

            // It doesn't matter if there is an active page or not
            // at this point it is considered resolved
            var app = cmdtarget.Application;
            var window = app.ActiveWindow;
            return new TargetSelection(window.Selection);
        }

        public IVisio.Selection Selection => this._item;

    }
}