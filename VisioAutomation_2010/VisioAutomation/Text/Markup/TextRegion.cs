namespace VisioAutomation.Text.Markup
{
    class TextRegion
    {
        // This class is used to identify continuos regions of text
        // mostly it is used to set character and paragraph formating

        // optionally a textregion may be associated with a text element
        public TextElement Element { get; set; }
        public Field Field { get; set; }
        public int Start { get; set; }
        public int Length { get; set; }

        public TextRegion()
        {
            // initialize an empty span
            this.Start = 0;
            this.Length = 0;

            // by default no text element is associated with this region
            this.Element = null;
        }

        public TextRegion(int start, TextElement el) :
            this()
        {
            this.Element = el;
            this.Start = start;
        }

        public int End
        {
            get { return (this.Start + this.Length); }
        }

    }
}