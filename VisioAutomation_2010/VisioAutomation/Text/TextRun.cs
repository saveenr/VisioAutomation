namespace VisioAutomation.Text
{
    public struct TextRun
    {
        private readonly int _begin;
        private readonly int _end;
        private readonly string _text;
        private readonly int _index;

        public int Begin => _begin;
        public int End => _end;
        public string Text => _text;
        public int Index => _index;

        internal TextRun(int index, int begin, int end, string text)
            : this()
        {
            this._index = index;
            this._begin = begin;
            this._end = end;
            this._text = text;
        }
    }
}