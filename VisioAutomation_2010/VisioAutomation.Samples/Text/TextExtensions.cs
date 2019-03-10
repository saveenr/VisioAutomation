using VA=VisioAutomation;

namespace VisioAutomationSamples
{
    public static class TextExtensions
    {
        public static VisioAutomation.Models.Text.Element AddElementEx(this VisioAutomation.Models.Text.Element p, string text,
                                                              int? font, double? size, int? color,
            VisioScripting.Models.AlignmentHorizontal? halign,
                                                              VA.Models.Text.CharStyle? cs)
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
                var c = new VisioAutomation.Models.Color.ColorRgb(color.Value);
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