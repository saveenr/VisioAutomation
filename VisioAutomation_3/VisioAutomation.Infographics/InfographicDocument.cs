using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Infographics
{
    public class InfographicDocument
    {
        public List<Block> Blocks;
        public InfographicDocument()
        {
            this.Blocks = new List<Block>();
        }

        public IVisio.Page RenderPage(IVisio.Document doc)
        {
            var pages = doc.Pages;
            var page = pages.Add();

            var pagesize = page.GetSize();

            var rendercontext = new RenderContext();
            rendercontext.CurrentUpperLeft = new VA.Drawing.Point(0, pagesize.Height);

            rendercontext.Page = page;
            foreach (var block in this.Blocks)
            {
                var blocksize = block.Render(rendercontext);
                rendercontext.CurrentUpperLeft = rendercontext.CurrentUpperLeft.Add(0.0, -blocksize.Height);
            }

            //page.ResizeToFitContents(0.5,0.5);

            return page;
        }
    }
}