using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetActiveSelection : TargetObject<IVisio.Selection>
    {

        public TargetActiveSelection() : base()
        {
        }

        public TargetActiveSelection(IVisio.Selection selection) : base(selection)
        {
        }

        public TargetActiveSelection Resolve(Client client)
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
            return new TargetActiveSelection(window.Selection);
        }

        public IVisio.Selection Selection => this._item;

    }
}