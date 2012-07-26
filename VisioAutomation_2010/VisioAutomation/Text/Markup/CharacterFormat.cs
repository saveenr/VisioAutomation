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

        private int? AsianFontID { get; set; }
        public int? Case { get; set; }
        public int? ComplexScriptFontID { get; set; }
        public double? ComplexScriptSize { get; set; }
        public bool? DoubleStrikeThrough{ get; set; }
        public bool? DoubleUnderline { get; set; }

        public int? LangID { get; set; }
        public int? Locale { get; set; }
        public int? LocalizeFont { get; set; }

        public bool? Overline { get; set; }
        public bool? Perpendicular { get; set; }

        public int? Pos { get; set; }
        public int? RTLText{ get; set; }
        public double? FontScalePercent { get; set; }
        public double? Letterspace{ get; set; }
        public bool? Strikethru { get; set; }
        public int? UseVertical { get; set; }

        public CharacterFormat()
        {
        }

        public void UpdateFrom(CharacterFormat other)
        {
            this.Color = other.Color;
            this.FontID = other.FontID;
            this.FontSizeInPoints = other.FontSizeInPoints;
            this.Style = other.Style;
            this.TransparencyPercent = other.TransparencyPercent;

            this.AsianFontID = other.AsianFontID;
            this.Case = other.Case;
            this.ComplexScriptFontID = other.ComplexScriptFontID;
            this.ComplexScriptSize = other.ComplexScriptSize;

            this.DoubleStrikeThrough = other.DoubleStrikeThrough;
            this.DoubleUnderline = other.DoubleUnderline;
            this.LangID= other.LangID;
            this.Locale = other.Locale;

            this.LocalizeFont = other.LocalizeFont;

            this.Overline = other.Overline;
            this.Perpendicular = other.Perpendicular;

            this.Pos = other.Pos;
            this.RTLText = other.RTLText;
            this.FontScalePercent = other.FontScalePercent;
            this.Letterspace = other.Letterspace;
            this.Strikethru = other.Strikethru;
            this.UseVertical = other.UseVertical;
        }

        public CharacterFormat Clone()
        {
            var fmt = new CharacterFormat();
            fmt.UpdateFrom(this);
            return fmt;
        }
    }
}