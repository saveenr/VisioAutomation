namespace VisioAutomation.Text.Markup
{
    public class TextRegion
    {
        public TextRegion()
        {
            this.Element = null;
            this.initialize_empty_span();
        }

        public TextRegion(TextElement el) :
            this()
        {
            this.Element = el;
        }

        public TextRegion(TextElement el, int start) :
            this()
        {
            this.Element = el;
            this.TextStartPos = start;
        }

        private void initialize_empty_span()
        {
            this.TextStartPos = 0;
            this.TextLength = 0;
        }

        public int TextEndPos
        {
            get { return (this.TextStartPos + this.TextLength); }
        }

        public TextElement Element { get; set; }

        public Field Field { get; set; }

        public int TextStartPos { get; set; }

        public int TextLength { get; set; }
    }
}