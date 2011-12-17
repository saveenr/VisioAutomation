namespace VisioAutomation.VDX.Elements
{
    public class TextRun
    {
        public int? CharacterFormatIndex { get; set; }
        public int? ParagraphFormatIndex { get; set; }
        public int? TabsFormatIndex { get; set; }

        public string Text { get; set; }

        public TextRun(string text)
        {
            this.Text = text;
        }

        public TextRun(string text, int? charfmt_index, int? parafmt_index, int? tabs_index)
        {
            this.Text = text;
            this.CharacterFormatIndex = charfmt_index;
            this.ParagraphFormatIndex = parafmt_index;
            this.TabsFormatIndex = tabs_index;
        }
    }
}