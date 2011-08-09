using Microsoft.Office.Interop.Visio;
using VisioAutomation.Drawing;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Infographics
{
    public abstract class Block
    {
        public abstract VA.Drawing.Size Render(RenderContext rc);
    }


    public class Header : Block
    {
        public string Text;

        public Header(string text)
        {
            this.Text = text;
        }

        public override Size Render(RenderContext rc)
        {
            var pagesize = rc.Page.GetSize();
            var size = new VA.Drawing.Size(pagesize.Width,1.0);
            var rect = DocUtil.BuildFromUpperLeft(rc.CurrentUpperLeft, size);

            var s = rc.Page.DrawRectangle(rect);
            if (this.Text != null)
            {
                s.Text = this.Text;                
            }

            return rect.Size;
        }
    }

    public static class DocUtil
    {
        public static VA.Drawing.Rectangle BuildFromUpperLeft(VA.Drawing.Point upperleft, VA.Drawing.Size s)
        {
            var rect = new VA.Drawing.Rectangle(upperleft.X, upperleft.Y - s.Height, upperleft.X + s.Width, upperleft.Y);
            return rect;
        }
    }
}