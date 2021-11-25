using VisioAutomation.Extensions;
using VA = VisioAutomation;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Dom
{
    public class Page : Node
    {
        public ShapeList Shapes { get; }
        public VisioAutomation.Core.Size? Size;
        public bool ResizeToFit;
        public VisioAutomation.Core.Size? ResizeToFitMargin;
        public Pages.PageFormatCells PageFormatCells;
        public Pages.PageLayoutCells PageLayoutCells;
        public string Name;
        public VisioAutomation.Models.LayoutStyles.LayoutStyleBase Layout;
        public IVisio.Page VisioPage;
        public RenderPerformanceSettings RenderPerformanceSettings { get; }

        public Page()
        {
            this.Shapes = new ShapeList();
            this.PageFormatCells = new Pages.PageFormatCells();
            this.PageLayoutCells = new VA.Pages.PageLayoutCells();

            this.RenderPerformanceSettings = new RenderPerformanceSettings();
            this.RenderPerformanceSettings.DeferRecalc = 0;
            
            // By Enable ScreenUpdating by default
            // If it is disabled it messes up page resizing (there may be a workaround)
            // TODO: Try the DrawTreeMultiNode2 unit test to see how setting it to 1 will affect the rendering

            this.RenderPerformanceSettings.ScreenUpdating = 1; 
            this.RenderPerformanceSettings.EnableAutoConnect = false;
            this.RenderPerformanceSettings.LiveDynamics = false;
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

            using (var perfscope = new RenderPerformanceScope(app, this.RenderPerformanceSettings))
            {
                if (this.Size.HasValue)
                {
                    this.PageFormatCells.Height = this.Size.Value.Height;
                    this.PageFormatCells.Width = this.Size.Value.Width;
                }

                var writer = new VASS.Writers.SidSrcWriter();
                writer.SetValues((short)page_sheet.ID, this.PageFormatCells);
                writer.SetValues((short)page_sheet.ID, this.PageLayoutCells);
                writer.Commit(page, ShapeSheet.CellValueType.Formula);
                
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