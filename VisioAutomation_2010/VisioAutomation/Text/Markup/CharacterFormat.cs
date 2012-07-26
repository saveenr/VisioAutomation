using VisioAutomation.Drawing;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;

namespace VisioAutomation.Text.Markup
{
    public class CharacterFormat
    {
        // http://msdn.microsoft.com/en-us/library/ff767069
        // this is used for dealing with character ranges typically

        public ColorRGB? Color { get; set; }
        public int? FontID { get; set; }
        public  double? FontSizeInPoints { get; set;  }
        public CharStyle? Style { get; set; }
        public int? TransparencyPercent { get; set; }
        public int? Case { get; set; }
        public bool? DoubleUnderline { get; set; }
        public int? LangID { get; set; }
        public bool? Overline { get; set; }
        public int? Pos { get; set; }
        public double? FontScalePercent { get; set; }
        public double? LetterspaceInPoints{ get; set; }
        public bool? Strikethru { get; set; }
        public int? UseVertical { get; set; }

        public CharacterFormat()
        {
        }

        public void UpdateFrom(CharacterFormat other)
        {
            this.FontID = other.FontID;
            this.Color = other.Color;
            this.Style = other.Style;
            this.Case = other.Case;
            this.Pos = other.Pos;
            this.FontSizeInPoints = other.FontSizeInPoints;
            this.TransparencyPercent = other.TransparencyPercent;
            this.DoubleUnderline = other.DoubleUnderline;
            this.FontScalePercent = other.FontScalePercent;
            this.LangID = other.LangID;
            this.LetterspaceInPoints = other.LetterspaceInPoints;
            this.Overline = other.Overline;
            this.Strikethru = other.Strikethru;
           
        }

        public CharacterFormat Clone()
        {
            var fmt = new CharacterFormat();
            fmt.UpdateFrom(this);
            return fmt;
        }

        public VA.Text.CharacterFormatCells ToCells()
        {
            var cells = new VA.Text.CharacterFormatCells();

            if (this.Case.HasValue)
            {
                cells.Case = this.Case.Value;
            }

            if (this.Color.HasValue)
            {
                cells.Color = this.Color.Value.ToFormula();
            }

            if (this.DoubleUnderline.HasValue)
            {
                cells.DoubleUnderline = this.DoubleUnderline.Value;
            }

            if (this.FontID != null)
            {
                cells.Font = this.FontID.Value;
            }

            if (this.FontScalePercent.HasValue)
            {
                cells.FontScale = this.FontScalePercent.Value / 100.0;
            }

            if (this.FontSizeInPoints.HasValue)
            {
                cells.Size = Convert.PointsToInches(this.FontSizeInPoints.Value);
            }

            if (this.LangID.HasValue)
            {
                cells.LangID = this.LangID.Value;
            }

            if (this.LetterspaceInPoints.HasValue)
            {
                cells.Letterspace = this.LetterspaceInPoints.Value / 100.0;
            }

            if (this.Overline.HasValue)
            {
                cells.Overline = this.Overline.Value;
            }

            if (this.Strikethru.HasValue)
            {
                cells.Strikethru = this.Strikethru.Value;
            }

            if (this.Style.HasValue)
            {
                cells.Style = (int)this.Style.Value;
            }

            if (this.TransparencyPercent.HasValue)
            {
                cells.Transparency = this.TransparencyPercent.Value / 100.0;
            }

            if (this.Pos.HasValue)
            {
                cells.Pos = this.Pos.Value;
            }

            return cells;
        }

    }
}