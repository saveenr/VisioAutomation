using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetActivePage
    {
        private IVisio.Page _page;

        public TargetActivePage()
        {
        }

        private TargetActivePage(IVisio.Page page)
        {
            this._page = page;
        }

        public IVisio.Page Page => this._page;

        public TargetActivePage Resolve(VisioScripting.Client client)
        {
            if (this._page != null)
            {
                return this;
            }
            var cmdtarget = client.GetCommandTargetPage();
            var page = cmdtarget.ActivePage;
            return new TargetActivePage(page);
        }

    }
}