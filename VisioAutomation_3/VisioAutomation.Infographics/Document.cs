using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Infographics
{
    public class Document
    {
        public List<Block> Blocks { get; private set; }
        public bool AutoResizePage { get; set; }
        public VA.Drawing.Size AutoResizeMargin { get; set; }

        private bool DeselectAfterDrawing=true;

        public Document()
        {
            this.Blocks = new List<Block>();
            this.AutoResizeMargin = new VA.Drawing.Size(0.5,0.5);
        }

        public IVisio.Page RenderPage(IVisio.Document doc)
        {
            var pages = doc.Pages;
            var page = pages.Add();

            var pagesize = page.GetSize();

            var rendercontext = new RenderContext();
            rendercontext.CurrentUpperLeft = new VA.Drawing.Point(0, pagesize.Height);
            rendercontext.PageWidth = pagesize.Width;

            rendercontext.Page = page;
            foreach (var block in this.Blocks)
            {
                var blocksize = block.Render(rendercontext);
                rendercontext.CurrentUpperLeft = rendercontext.CurrentUpperLeft.Add(0.0, -blocksize.Height);
            }

            if (this.AutoResizePage)
            {
                page.ResizeToFitContents(this.AutoResizeMargin);
            }

            if (this.DeselectAfterDrawing)
            {
                var app = page.Application;
                var window = app.ActiveWindow;
                window.DeselectAll();
            }

            return page;
        }
    }
}