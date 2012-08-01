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
            var el = p.AddElement(text);
            if (font != null)
            {
                el.CharacterFormatCells.Font = font.Value;
            }
            if (size.HasValue)
            {
                el.CharacterFormatCells.Size = VA.Convert.PointsToInches(size.Value);
            }
            if (color.HasValue)
            {
                var c = new VA.Drawing.ColorRGB(color.Value);
                el.CharacterFormatCells.Color = c.ToFormula();
            }
            if (halign.HasValue)
            {
                el.ParagraphFormatCells.HorizontalAlign = (int) halign.Value;
            }

            if (cs.HasValue)
            {
                el.CharacterFormatCells.Style = (int) cs;
            }

            return el;
        }

    }
}