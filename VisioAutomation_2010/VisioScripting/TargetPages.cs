using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting
{
    public class TargetPages : TargetObjects<IVisio.Page>
    {

        private TargetPages() : base()
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

            if (this.Resolved)
            {
                return this;
            }

            // Otherwise perform resolution
            // Try to use the active page as the default target for the operation

            var cmdtarget = client.GetCommandTargetPage();

            if (cmdtarget.ActivePage == null)
            {
                throw new VisioAutomation.Exceptions.AutomationException("Resolving failed: No active page available");
            }

            client.Output.WriteVerbose("Resolving to active page (name={0})", cmdtarget.ActivePage.Name);

            return new TargetPages(cmdtarget.ActivePage);
        }

        public IList<IVisio.Page> Pages => this._get_items_safe();

        public static TargetPages Auto => new TargetPages();
    }
}