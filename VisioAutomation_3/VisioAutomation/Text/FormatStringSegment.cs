namespace VisioAutomation
{
    internal class FormatStringSegment
    {
        public string Text { get; private set; }
        public int Index { get; private set; }
        public int Start { get; private set; }
        public int End { get; private set; }

        public FormatStringSegment(string text, int index, int start, int end)
        {
            this.Text = text;
            this.Index = index;
            this.Start = start;
            this.End = end;
        }
    }
}