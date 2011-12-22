using VisioAutomation.Drawing;
using VA = VisioAutomation;

namespace VisioAutomation.Text.Markup
{
    public class ParagraphFormat
    {
        public AlignmentHorizontal? HAlign { get; set; }
        public double? Indent { get; set; }
        public bool? Bullets { get; set; }

        public void UpdateFrom(ParagraphFormat other)
        {
            this.HAlign =  other.HAlign;
            this.Indent =  other.Indent;
            this.Bullets = other.Bullets;
        }

        public ParagraphFormat Duplicate()
        {
            var fmt = new ParagraphFormat();
            fmt.UpdateFrom(this);
            return fmt;
        }
    }

}