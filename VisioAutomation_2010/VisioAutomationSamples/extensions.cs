using VA=VisioAutomation;

namespace VisioAutomationSamples
{
    public static class extensions
    {
        public static VA.Text.Markup.TextElement AddElementEx(this VA.Text.Markup.TextElement p, string text,
                                                              int? font, double? size, int? color,
                                                              VA.Drawing.AlignmentHorizontal? halign,
                                                              VA.Text.CharStyle? cs)
        {
            var el = p.AppendElement(text);
            if (font != null)
            {
                el.CharacterFormat.Font = font;
            }
            if (size.HasValue)
            {
                el.CharacterFormat.Size = size.Value;
            }
            if (color.HasValue)
            {
                el.CharacterFormat.Color = new VA.Drawing.ColorRGB(color.Value);
            }
            if (halign.HasValue)
            {
                el.ParagraphFormat.HAlign = halign.Value;
            }

            if (cs.HasValue)
            {
                el.CharacterFormat.Style = cs;
            }

            return el;
        }

    }
}