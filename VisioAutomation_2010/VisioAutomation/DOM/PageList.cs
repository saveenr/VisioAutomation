using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections;
using System.Collections.Generic;

namespace VisioAutomation.DOM
{
    public class PageList : Node, IEnumerable<Page>
    {
        private NodeList<Page> pages;

        public PageList()
        {
            this.pages = new NodeList<Page>(this);
        }

        public IEnumerator<Page> GetEnumerator()
        {
            foreach (var i in this.pages)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     
        {                                           
            return GetEnumerator();
        }

        public void Add(Page page)
        {
            this.pages.Add(page);
        }

        public int Count
        {
            get { return this.pages.Count; }
        }

        public IList<IVisio.Page> Render(IVisio.Document doc)
        {
            var vpages = new List<IVisio.Page>(this.Count);
            foreach (var dompage in this.pages)
            {
                var vpage = dompage.Render(doc);
                vpages.Add(vpage);
            }
            return vpages;
        }

        public IList<IVisio.Page> Render(IVisio.Page startpage)
        {
            var doc = startpage.Document;
            int count = 0;
            var vpages = new List<IVisio.Page>(this.Count);
            foreach (var dompage in this.pages)
            {
                if (count == 0)
                {
                    dompage.Render(startpage);
                    vpages.Add(startpage);
                }
                else
                {
                    var vpage = dompage.Render(doc);
                    vpages.Add(vpage);
                }
            }
            return vpages;
        }

    }
}