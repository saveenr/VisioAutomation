using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public struct TextRun
    {
        public int Begin { get; private set; }
        public int End { get; private set; }
        public string Text { get; private set; }
        public int Index { get; private set; }

        internal TextRun(int index, int begin, int end, string text)
            : this()
        {
            this.Index = index;
            this.Begin = begin;
            this.End = end;
            this.Text = text;
        }
        
        public override string ToString()
        {
            var s = string.Format("(Begin={0},End={1})", this.Begin, this.End);
            return s;
        }
    }
}