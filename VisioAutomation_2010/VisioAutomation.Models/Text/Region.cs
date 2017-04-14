namespace VisioAutomation.Models.Text
{
    class Region
    {
        // This class is used to identify continuos regions of text
        // mostly it is used to set character and paragraph formating

        // optionally a textregion may be associated with a text element
        public Element Element { get; internal set; }
        public Field Field { get; internal set; }
        public int Start { get; internal set; }
        public int Length { get; internal set; }

        internal Region()
        {
            // initialize an empty span
            this.Start = 0;
            this.Length = 0;

            // by default no text element is associated with this region
            this.Element = null;

            // by default no field is associated with this region
            this.Field = null;

        }

        internal Region(int start, Element el) :
            this()
        {
            this.Element = el;
            this.Start = start;
        }

        internal Region(int start, Field field) :
            this()
        {
            this.Field = field;
            this.Start = start;
            this.Length = field.PlaceholderText.Length;
        }

        public int End
        {
            get { return (this.Start + this.Length); }
        }

    }
}