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
        public VA.Application.PerfSettings PerfSettings { get; private set; }

        public Page()
        {
            this.Shapes = new ShapeList();
            this.PageCells = new VA.Pages.PageCells();

            this.PerfSettings = new VA.Application.PerfSettings();
            this.PerfSettings.DeferRecalc = 0;
            
            
            // By Enable ScreenUpdating by default
            // If it is disabled it messes up page resizing (there may be a workaround)
            // TODO: Try the DrawTreeMultiNode2 unit test to see how setting it to 1 will affect the rendering

            this.PerfSettings.ScreenUpdating = 1; 
            this.PerfSettings.EnableAutoConnect = false;
            this.PerfSettings.LiveDynamics = false;
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

            var app = page.Application;


            using (var perfscope = new VA.Application.PerfScope(app, PerfSettings))
            {
                if (this.Size.HasValue)
                {
                    this.PageCells.PageHeight = this.Size.Value.Height;
                    this.PageCells.PageWidth = this.Size.Value.Width;
                }

                var update = new VA.ShapeSheet.Update();
                this.PageCells.Apply(update, (short)page_sheet.ID);
                update.Execute(page);


                // Then render the shapes
                this.Shapes.Render(page);

                // Perform any additional layout
                if (this.Layout != null)
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
}