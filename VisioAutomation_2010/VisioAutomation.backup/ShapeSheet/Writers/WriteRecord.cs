namespace VisioAutomation.ShapeSheet.Writers
{
    internal struct WriteRecord
    {
        public readonly SidSrc SidSrc;
        public readonly string Value;
        public WriteRecord(SidSrc sidsrc, string value)
        {
            this.SidSrc = sidsrc;
            this.Value = value;
        }

        public WriteRecord(Src src, string value)
        {
            this.SidSrc = new SidSrc(-1,src);
            this.Value = value;
        }
    }
}