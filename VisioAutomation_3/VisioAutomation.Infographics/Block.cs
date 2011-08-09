using Microsoft.Office.Interop.Visio;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Infographics
{
    public abstract class Block
    {
        public abstract VA.Drawing.Size Render(RenderContext rc);
    }
}