using VisioAutomation.Drawing;
using VisioAutomation.Text;
using VisioAutomation.Text.Markup;

namespace VisioAutomationSamples
{
    public static class extensions
    {
        public static TextElement AddElementEx(this TextElement p, string text,
                                                              string font, double? size, int? color,
                                                              AlignmentHorizontal? halign,
                                                              CharStyle? cs)
        {
            var el = p.AppendElement(text);
            if (font != null)
            {
                el.CharacterFormat.Font = font;
            }
            if (size.HasValue)
            {
                el.CharacterFormat.FontSize = size.Value;
            }
            if (color.HasValue)
            {
                el.CharacterFormat.Color = new ColorRGB(color.Value);
            }
            if (halign.HasValue)
            {
                el.ParagraphFormat.HAlign = halign.Value;
            }

            if (cs.HasValue)
            {
                el.CharacterFormat.CharStyle = cs;
            }

            return el;
        }

    }
}