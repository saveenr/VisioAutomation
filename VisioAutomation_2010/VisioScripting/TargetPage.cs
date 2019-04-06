using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetPage : TargetObject<IVisio.Page>
    {

        public TargetPage() : base()
        {
        }

        public TargetPage(Microsoft.Office.Interop.Visio.Page page) : base(page)
        {
        }

        public TargetPage(Microsoft.Office.Interop.Visio.Page page, bool isresolved) : base(page, isresolved)
        {
        }

        public TargetPage Resolve(VisioScripting.Client client)
        {
            if (!this.IsResolved)
            {
                var cmdtarget = client.GetCommandTargetPage();

                // It doesn't matter if there is an active page or not
                // at this point it is considered resolved
                return new TargetPage(cmdtarget.ActivePage, true);
            }
            else
            {
                return this;
            }
        }

        public IVisio.Page Page => this._item;
    }
}