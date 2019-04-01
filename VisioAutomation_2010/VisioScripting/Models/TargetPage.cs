using VisioAutomation.Exceptions;

namespace VisioScripting.Models
{
    public class TargetPage
    {
        public Microsoft.Office.Interop.Visio.Page Page { get; private set; }

        public TargetPage()
        {
            // This explicitly means that the active document will be used
            this.Page = null;
        }

        public TargetPage(Microsoft.Office.Interop.Visio.Page page)
        {
            this.Page = page;
        }

        public Microsoft.Office.Interop.Visio.Page Resolve(VisioScripting.Client client)
        {
            if (this.Page == null)
            {
                var cmdtarget = client.GetCommandTargetPage();
                this.Page = cmdtarget.ActivePage;
            }

            if (this.Page == null)
            {
                throw new VisioOperationException("Unvalid State No Pages");
            }

            return this.Page;
        }
    }
}