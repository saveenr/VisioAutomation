using VisioAutomation.Colors;
using VA=VisioAutomation;

namespace VisioAutomationSamples
{
    public static class TextExtensions
    {
        public static VisioAutomation.Models.Text.TextElement AddElementEx(this VisioAutomation.Models.Text.TextElement p, string text,
                                                              int? font, double? size, int? color,
                                                              VA.Drawing.AlignmentHorizontal? halign,
                                                              VA.Text.CharStyle? cs)
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
                var c = new ColorRGB(color.Value);
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