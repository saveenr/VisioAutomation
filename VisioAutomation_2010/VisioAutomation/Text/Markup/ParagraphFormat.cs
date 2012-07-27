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

        public VA.Text.ParagraphFormatCells ToCells()
        {

            var paracells = new VA.Text.ParagraphFormatCells();

            // Handle bullets
            if (this.Bullets.HasValue && this.Bullets.Value)
            {
                const int bullet_type = 1;
                const int base_indent_size = 25;
                int indent_first = -base_indent_size;
                int indent_left = base_indent_size;

                paracells.IndentFirst = indent_first;
                paracells.IndentLeft = indent_left;
                paracells.Bullet = bullet_type;
            }

            if (this.IndentFirstInPoints.HasValue)
            {
                paracells.IndentFirst = VA.Convert.PointsToInches(this.IndentFirstInPoints.Value);
            }

            if (this.IndentLeftInPoints.HasValue)
            {
                paracells.IndentLeft = VA.Convert.PointsToInches(this.IndentLeftInPoints.Value);
            }

            if (this.HAlign.HasValue)
            {
                paracells.HorizontalAlign = (int)this.HAlign;
            }

            return paracells;
        }
    }

}