using VisioAutomation.Extensions;
using VisioAutomation.Models.LayoutStyles;
using VisioAutomation.Pages;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Dom
{
    public class Page : Node
    {
        public ShapeList Shapes { get; }
        public VisioAutomation.Geometry.Size? Size;
        public bool ResizeToFit;
        public VisioAutomation.Geometry.Size? ResizeToFitMargin;
        public Pages.PageFormatCells PageFormatCells;
        public Pages.PageLayoutCells PageLayoutCells;
        public string Name;
        public VisioAutomation.Models.LayoutStyles.LayoutStyleBase Layout;
        public IVisio.Page VisioPage;
        public RenderPerforfmanceSettings RenderPerforfmanceSettings { get; }

        public Page()
        {
            this.Shapes = new ShapeList();
            this.PageFormatCells = new Pages.PageFormatCells();
            this.PageLayoutCells = new PageLayoutCells();

            this.RenderPerforfmanceSettings = new RenderPerforfmanceSettings();
            this.RenderPerforfmanceSettings.DeferRecalc = 0;
            
            // By Enable ScreenUpdating by default
            // If it is disabled it messes up page resizing (there may be a workaround)
            // TODO: Try the DrawTreeMultiNode2 unit test to see how setting it to 1 will affect the rendering

            this.RenderPerforfmanceSettings.ScreenUpdating = 1; 
            this.RenderPerforfmanceSettings.EnableAutoConnect = false;
            this.RenderPerforfmanceSettings.LiveDynamics = false;
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

            using (var perfscope = new RenderPerformanceScope(app, this.RenderPerforfmanceSettings))
            {
                if (this.Size.HasValue)
                {
                    this.PageFormatCells.Height = this.Size.Value.Height;
                    this.PageFormatCells.Width = this.Size.Value.Width;
                }

                var writer = new SidSrcWriter();
                writer.SetFormulas((short)page_sheet.ID, this.PageFormatCells);
                writer.SetFormulas((short)page_sheet.ID, this.PageLayoutCells);
                writer.Commit(page);
                
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