using VisioAutomation.Drawing;
using VA = VisioAutomation;

namespace VisioAutomation.Text.Markup
{
    public class ParagraphFormat
    {
        // http://msdn.microsoft.com/en-us/library/ff767385

        public AlignmentHorizontal? HAlign { get; set; }
        public double? IndentFirstInPoints { get; set; }
        public double? IndentLeftInPoints { get; set; }
        public bool? Bullets { get; set; }

        public void UpdateFrom(ParagraphFormat other)
        {
            this.HAlign =  other.HAlign;
            this.IndentFirstInPoints = other.IndentFirstInPoints;
            this.IndentLeftInPoints = other.IndentLeftInPoints;
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