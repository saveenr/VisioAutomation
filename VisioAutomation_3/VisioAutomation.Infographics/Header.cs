using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Infographics
{
    public class Header : Block
    {
        public string Text;

        public Header(string text)
        {
            this.Text = text;
        }

        public override VA.Drawing.Size Render(RenderContext rc)
        {
            var pagesize = rc.Page.GetSize();
            var size = new VA.Drawing.Size(pagesize.Width, 1.0);
            var rect = DocUtil.BuildFromUpperLeft(rc.CurrentUpperLeft, size);

            var s = rc.Page.DrawRectangle(rect);
            if (this.Text != null)
            {
                s.Text = this.Text;                
            }

            return rect.Size;
        }
    }
}