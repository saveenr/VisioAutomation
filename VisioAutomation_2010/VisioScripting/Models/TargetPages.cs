using System.Collections.Generic;
using VisioAutomation.Exceptions;

namespace VisioScripting.Models
{
    public class TargetPages
    {
        public IList<Microsoft.Office.Interop.Visio.Page> Pages { get; private set; }

        public TargetPages()
        {
            // This explicitly means that the active document will be used
            this.Pages = null;
        }

        public TargetPages(IList<Microsoft.Office.Interop.Visio.Page> pages)
        {
            this.Pages = pages;
        }

        public TargetPages( params Microsoft.Office.Interop.Visio.Page[] pages)
        {
            this.Pages = pages;
        }


        public IList<Microsoft.Office.Interop.Visio.Page> Resolve(VisioScripting.Client client)
        {
            if (this.Pages == null)
            {
                var cmdtarget = client.GetCommandTargetPage();
                this.Pages = new List<Microsoft.Office.Interop.Visio.Page> {cmdtarget.ActivePage};
            }

            if (this.Pages == null)
            {
                throw new VisioOperationException("Unvalid State No Pages");
            }

            if (this.Pages.Count < 1)
            {
                throw new VisioOperationException("Unvalid State No Pages");
            }

            return this.Pages;
        }
    }

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