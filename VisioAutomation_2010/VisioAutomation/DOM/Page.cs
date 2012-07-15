using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;
using System.Collections;
using System.Collections.Generic;

namespace VisioAutomation.DOM
{
    public class Document
    {
        public PageList Pages;

        public Document()
        {
            this.Pages = new PageList();
        }

        public IVisio.Document Render(IVisio.Application app)
        {
            var appdocs = app.Documents;
            var vdoc = appdocs.Add("");
            var docpages = vdoc.Pages;
            var starpage = docpages[1];
            this.Pages.Render(starpage);
            return vdoc;
        }
    }

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

        IEnumerator IEnumerable.GetEnumerator()     // Explicit implementation
        {                                           // keeps it hidden.
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
                if (count==0)
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


    public class Page : Node
    {
        public ShapeList Shapes { get; private set; }
        public VA.Drawing.Size? Size;
        public bool ResizeToFit;
        public VA.Drawing.Size? ResizeToFitMargin;
        public VA.Pages.PageCells PageCells;

        public Page()
        {
            this.Shapes = new ShapeList();
            this.PageCells = new VA.Pages.PageCells();
        }

        public IVisio.Page Render(IVisio.Document doc)
        {
            if (doc== null)
            {
                throw new System.ArgumentNullException("doc");
            }

            var pages = doc.Pages;
            var page = pages.Add();

            this.Render(page);
            
            return page;
        }

        public void Render(IVisio.Page page)
        {
            if (page == null)
            {
                throw new System.ArgumentNullException("page");
            }

            // First handle any page properties
            var page_sheet = page.PageSheet;
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            this.PageCells.Apply(update, (short)page_sheet.ID);
            update.Execute(page);

            if (this.Size.HasValue)
            {
                page.SetSize(this.Size.Value);
            }
            
            // Then render the shapes
            this.Shapes.Render(page);

            // Optionally, perform page resizing to fit contents
            if (this.ResizeToFit)
            {
                if (this.ResizeToFitMargin.HasValue)
                {
                    page.ResizeToFitContents(this.ResizeToFitMargin.Value);
                }
                else
                {
                    page.ResizeToFitContents();
                }
            }
        }
    }
}