namespace VisioAutomation.ShapeSheet.Writers
{
    internal struct WriteRecord
    {
        public readonly Core.SidSrc SidSrc;
        public readonly string Value;

        public WriteRecord(Core.SidSrc sidsrc, string value)
        {
            this.SidSrc = sidsrc;
            this.Value = value;
        }

        public WriteRecord(Core.Src src, string value)
        {
            this.SidSrc = new Core.SidSrc(-1, src);
            this.Value = value;
        }
    }
}