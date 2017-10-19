using System.Collections.Generic;
using VisioAutomation.Exceptions;
using VisioScripting.Commands;

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
                var cmdtarget = client.GetCommandTarget(CommandTargetFlags.Application | CommandTargetFlags.ActiveDocument |
                                                        CommandTargetFlags.ActivePage);

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
}