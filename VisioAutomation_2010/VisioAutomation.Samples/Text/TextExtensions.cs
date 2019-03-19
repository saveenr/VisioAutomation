using VA=VisioAutomation;
using VAM = VisioAutomation.Models;

namespace VisioAutomationSamples
{
    public static class TextExtensions
    {
        public static VAM.Text.Element AddElementEx(
            this VAM.Text.Element p, 
            string text,
            int? font, double? size, int? color,
            VisioScripting.Models.AlignmentHorizontal? halign,
            VAM.Text.CharStyle? cs)
        {
            var el = p.AddElement(text);

            if (font != null)
            {
                el.CharacterFormatting.Font = font.Value;
            }

            if (size.HasValue)
            {
                el.CharacterFormatting.Size = string.Format("{0}pt",size.Value);
            }

            if (color.HasValue)
            {
                var c = new VAM.Color.ColorRgb(color.Value);
                el.CharacterFormatting.Color = c.ToFormula();
            }

            if (halign.HasValue)
            {
                el.ParagraphFormatting.HorizontalAlign = (int) halign.Value;
            }

            if (cs.HasValue)
            {
                el.CharacterFormatting.Style = (int) cs;
            }

            return el;
        }

    }
}