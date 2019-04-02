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
            if (this.Items == null)
            {
                var cmdtarget = client.GetCommandTargetPage();
                if (cmdtarget.ActivePage != null)
                {
                    var pages = new List<IVisio.Page> { cmdtarget.ActivePage };
                    return new TargetPages(pages);
                }
                else
                {
                    return new TargetPages(new List<IVisio.Page>(0));
                }
            }
            else
            {
                return this;
            }
        }
    }
}