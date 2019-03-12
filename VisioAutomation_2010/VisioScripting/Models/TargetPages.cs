using System.Collections.Generic;
using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
{
    public class TargetPages
    {
        public IList<IVisio.Page> Pages { get; private set; }

        public TargetPages()
        {
            // This explicitly means that the active document will be used
            this.Pages = null;
        }

        public TargetPages(IList<IVisio.Page> pages)
        {
            this.Pages = pages;
        }

        public TargetPages( params IVisio.Page[] pages)
        {
            this.Pages = pages;
        }


        public IList<IVisio.Page> Resolve(VisioScripting.Client client)
        {
            if (this.Pages == null)
            {
                var cmdtarget = client.GetCommandTargetPage();
                this.Pages = new List<IVisio.Page> {cmdtarget.ActivePage};
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
        public IVisio.Page Page { get; private set; }

        public TargetPage()
        {
            // This explicitly means that the active document will be used
            this.Page = null;
        }

        public TargetPage(IVisio.Page page)
        {
            this.Page = page;
        }

        public IVisio.Page Resolve(VisioScripting.Client client)
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