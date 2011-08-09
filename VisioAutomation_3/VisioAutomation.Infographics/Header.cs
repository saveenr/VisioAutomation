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

            var s = rc.Page.DrawRectangle(rect);
            if (this.Text != null)
            {
                s.Text = this.Text;                
            }

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            var charfmt = new VA.Text.CharacterFormatCells();
            charfmt.Size = VA.Convert.PointsToInches(this.FontSize);
            charfmt.Font = rc.GetFontID(rc.DefaultFont);

            VA.Text.CharStyle cs = 0;

            if (this.Bold)
            {
                cs |= VA.Text.CharStyle.Bold;
            }

            charfmt.Style = (int) cs;

            var parafmt = new VA.Text.ParagraphFormatCells();
            parafmt.HorizontalAlign = 0;

            var bkfmt = rc.GetDefaultBkfmt();

            var s_id = s.ID16;
            charfmt.Apply(update,s_id,0);
            bkfmt.Apply(update,s_id);
            parafmt.Apply(update,s_id,0);

            update.Execute(rc.Page);
            return rect.Size;
        }
    }
}