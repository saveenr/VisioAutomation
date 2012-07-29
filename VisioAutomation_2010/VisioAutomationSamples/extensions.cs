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
            var el = p.Add(text);
            if (font != null)
            {
                el.CharacterFormat.Font = font.Value;
            }
            if (size.HasValue)
            {
                el.CharacterFormat.Size = VA.Convert.PointsToInches(size.Value);
            }
            if (color.HasValue)
            {
                var c = new VA.Drawing.ColorRGB(color.Value);
                el.CharacterFormat.Color = c.ToFormula();
            }
            if (halign.HasValue)
            {
                el.ParagraphFormat.HAlign = halign.Value;
            }

            if (cs.HasValue)
            {
                el.CharacterFormat.Style = (int) cs;
            }

            return el;
        }

    }
}