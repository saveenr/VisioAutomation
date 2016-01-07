using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections;
using System.Collections.Generic;

namespace VisioAutomation.DOM
{
    public class PageList : Node, IEnumerable<Page>
    {
        private readonly NodeList<Page> pagenodes;

        public PageList()
        {
            this.pagenodes = new NodeList<Page>(this);
        }

        public IEnumerator<Page> GetEnumerator()
        {
            foreach (var i in this.pagenodes)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     
        {                                           
            return this.GetEnumerator();
        }

        public void Add(Page page)
        {
            this.pagenodes.Add(page);
        }

        public int Count
        {
            get { return this.pagenodes.Count; }
        }

        public IList<IVisio.Page> Render(IVisio.Document doc)
        {
            var pages = new List<IVisio.Page>(this.Count);
            foreach (var pagenode in this.pagenodes)
            {
                var page = pagenode.Render(doc);
                pages.Add(page);
            }
            return pages;
        }

        public IList<IVisio.Page> Render(IVisio.Page startpage)
        {
            var doc = startpage.Document;
            int count = 0;
            var pages = new List<IVisio.Page>(this.Count);

            var app = doc.Application;
            var active_window = app.ActiveWindow;
            foreach (var pagenode in this.pagenodes)
            {
                if (count == 0)
                {
                    pagenode.Render(startpage);
                    pages.Add(startpage);
                }
                else
                {
                    var rendered_page = pagenode.Render(doc);
                    pages.Add(rendered_page);
                }

                active_window.ViewFit = 1; // 1==visFitPage - adjust the zoom
                count++;
            }
            return pages;
        }

    }
}