using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public struct TextRun
    {
        public int Begin { get; private set; }
        public int End { get; private set; }
        public string Text { get; private set; }
        public int Index { get; private set; }

        public TextRun(int index, int begin, int end, string text)
            : this()
        {
            this.Index = index;
            this.Begin = begin;
            this.End = end;
            this.Text = text;
        }
        
        public override string ToString()
        {
            string t = this.Text ?? string.Empty;
            var s = string.Format("(Begin={0},End={1},Text=\"{2}\")", this.Begin, this.End, t);
            return s;
        }
    }
}