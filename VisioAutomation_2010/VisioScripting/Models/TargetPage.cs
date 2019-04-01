using VisioAutomation.Exceptions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioScripting.Models
{
    public class TargetPage : TargetObject<IVisio.Page>
    {

        public TargetPage() : base()
        {
        }

        public TargetPage(Microsoft.Office.Interop.Visio.Page page) : base (page)
        {
        }

        public TargetPage(Microsoft.Office.Interop.Visio.Page page, bool isresolved) : base (page,isresolved)
        {
        }

        public TargetPage Resolve(VisioScripting.Client client)
        {
            if (!this.IsResolved)
            {
                var cmdtarget = client.GetCommandTargetPage();
                if (cmdtarget.ActivePage != null)
                {
                    return new TargetPage(cmdtarget.ActivePage);
                }
                else
                {
                    return new TargetPage(cmdtarget.ActivePage, true);
                }
            }
            else
            {
                return this;
            }
        }
    }
}