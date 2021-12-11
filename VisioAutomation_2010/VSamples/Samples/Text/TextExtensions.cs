using VAM = VisioAutomation.Models;

namespace VSamples.Samples.Text
{
    public class TextFmt
    {
        public int? FontID;
        public double? FontSize;
        public int? Color;
        public VisioScripting.Models.AlignmentHorizontal? HAlign;
        public VAM.Text.CharStyle? CharStyle;
    }

    public static class TextExtensions
    {
        public static VAM.Text.Element AddElementEx(
            this VAM.Text.Element p, 
            string text,
            TextFmt textfmt)
        {
            var el = p.AddElement(text);

            if (textfmt !=null && textfmt.FontID != null)
            {
                el.CharacterFormatting.Font = textfmt.FontID.Value;
            }

            if (textfmt != null && textfmt.FontSize.HasValue)
            {
                el.CharacterFormatting.Size = string.Format("{0}pt", textfmt.FontSize.Value);
            }

            if (textfmt != null && textfmt.Color.HasValue)
            {
                var c = new VAM.Color.ColorRgb(textfmt.Color.Value);
                el.CharacterFormatting.Color = c.ToFormula();
            }

            if (textfmt != null && textfmt.HAlign.HasValue)
            {
                el.ParagraphFormatting.HorizontalAlign = (int)textfmt.HAlign.Value;
            }

            if (textfmt != null && textfmt.CharStyle.HasValue)
            {
                el.CharacterFormatting.Style = (int)textfmt.CharStyle;
            }

            return el;
        }

    }
}