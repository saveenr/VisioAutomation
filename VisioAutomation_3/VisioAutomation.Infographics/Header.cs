using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace VisioAutomation.Infographics
{
    public class Header : Block
    {
        public string Text;
        public double FontSize = 18.0;
        public double Margin = 0.25;
        public bool Bold = false;
        public Header(string text)
        {
            this.Text = text;
        }

        public override VA.Drawing.Size Render(RenderContext rc)
        {
            double height = VA.Convert.PointsToInches(this.FontSize) + (2*this.Margin);
            double nearest = 1.0/4.0;
            height = System.Math.Round(height/nearest, System.MidpointRounding.AwayFromZero)*nearest;
            var size = new VA.Drawing.Size(rc.PageWidth, height);
            var rect = DocUtil.BuildFromUpperLeft(rc.CurrentUpperLeft, size);


            var xdoc = new VA.DOM.Document();
            var tile = xdoc.DrawRectangle(rect);
            tile.Text = this.Text;

            tile.ShapeCells.CharSize = VA.Convert.PointsToInches(this.FontSize);
            tile.ShapeCells.CharFont = rc.GetFontID(rc.DefaultFont);

            VA.Text.CharStyle cs = 0;

            if (this.Bold)
            {
                cs |= VA.Text.CharStyle.Bold;
            }

            tile.ShapeCells.CharStyle = (int) cs;
            tile.ShapeCells.HAlign = 0;
            tile.ShapeCells.FillForegnd = rc.BKColor.ToFormula();
            tile.ShapeCells.LinePattern = 0;
            tile.ShapeCells.LineWeight = 0;
            tile.ShapeCells.LineColor = 0;

            xdoc.Render(rc.Page);

            return rect.Size;
        }
    }
}