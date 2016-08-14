using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomation.DOM
{
    public class Page : Node
    {
        public ShapeList Shapes { get; }
        public Drawing.Size? Size;
        public bool ResizeToFit;
        public Drawing.Size? ResizeToFitMargin;
        public Pages.PageCells PageCells;
        public string Name;
        public Pages.PageLayout.Layout Layout;
        public IVisio.Page VisioPage;
        public Application.PerfSettings PerfSettings { get; }

        public Page()
        {
            this.Shapes = new ShapeList();
            this.PageCells = new Pages.PageCells();

            this.PerfSettings = new Application.PerfSettings();
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
                throw new System.ArgumentNullException(nameof(doc));
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
                throw new System.ArgumentNullException(nameof(page));
            }


            // First handle any page properties
            if (this.Name!=null)
            {
                page.NameU = this.Name;
            }

            this.VisioPage = page;

            var page_sheet = page.PageSheet;

            var app = page.Application;


            using (var perfscope = new Application.PerfScope(app, this.PerfSettings))
            {
                if (this.Size.HasValue)
                {
                    this.PageCells.PageHeight = this.Size.Value.Height;
                    this.PageCells.PageWidth = this.Size.Value.Width;
                }

                var update = new FormulaWriterSIDSRC();
                this.PageCells.SetFormulas((short)page_sheet.ID, update);
                update.Commit(page);
                
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