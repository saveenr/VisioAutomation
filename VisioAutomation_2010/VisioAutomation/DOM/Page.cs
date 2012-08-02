using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.DOM
{
    public class Page : Node
    {
        public ShapeList Shapes { get; private set; }
        public VA.Drawing.Size? Size;
        public bool ResizeToFit;
        public VA.Drawing.Size? ResizeToFitMargin;
        public VA.Pages.PageCells PageCells;
        public string Name;
        public VA.Layout.PageLayout.Layout Layout;
        public IVisio.Page VisioPage;

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
            this.VisioPage = page;

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
            if (this.Name!=null)
            {
                page.NameU = this.Name;
            }

            this.VisioPage = page;

            var page_sheet = page.PageSheet;
            
            var update = new VA.ShapeSheet.Update.UpdateBase();
            this.PageCells.Apply(update, (short)page_sheet.ID);
            update.Execute(page);

            if (this.Size.HasValue)
            {
                page.SetSize(this.Size.Value);
            }
            
            // Then render the shapes
            this.Shapes.Render(page);

            // Perform any additional layout
            if (this.Layout!=null)
            {
                this.Layout.Apply(page);
            }

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