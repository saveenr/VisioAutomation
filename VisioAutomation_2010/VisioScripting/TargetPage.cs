using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetPage : TargetObject<IVisio.Page>
    {

        public TargetPage() : base()
        {
        }

        public TargetPage(IVisio.Page page) : base(page)
        {
        }

        public TargetPage Resolve(Client client)
        {
            if (this.Resolved)
            {
                return this;
            }

            var cmdtarget = client.GetCommandTargetPage();

            client.Output.WriteVerbose("Resolving to active page (name={0})", cmdtarget.ActivePage.Name);

            return new TargetPage(cmdtarget.ActivePage);
        }

        public IVisio.Page Page => this._get_item_safe();

        public static TargetPage Active => new TargetPage();
    }
}