using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetPages : TargetObjects<IVisio.Page>
    {

        public TargetPages() : base()
        {
        }

        public TargetPages(IList<IVisio.Page> pages) : base (pages)
        {
        }

        public TargetPages( params IVisio.Page[] pages) : base (pages)
        {

        }


        public TargetPages Resolve(VisioScripting.Client client)
        {
            // Handle the case where the object is already resolved
            if (this._items != null)
            {
                return this;
            }

            // Otherwise perform resolution
            // Try to use the active page as the default target for the operation

            var cmdtarget = client.GetCommandTargetPage();

            if (cmdtarget.ActivePage == null)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            return new TargetPages(cmdtarget.ActivePage);
        }

        public IList<IVisio.Page> Pages => this._items;
    }
}