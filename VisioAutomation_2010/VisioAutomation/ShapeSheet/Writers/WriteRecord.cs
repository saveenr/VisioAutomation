namespace VisioAutomation.ShapeSheet.Writers
{
    internal struct WriteRecord
    {
        public readonly VisioAutomation.Core.SidSrc SidSrc;
        public readonly string Value;
        public WriteRecord(VisioAutomation.Core.SidSrc sidsrc, string value)
        {
            this.SidSrc = sidsrc;
            this.Value = value;
        }

        public WriteRecord(VisioAutomation.Core.Src src, string value)
        {
            this.SidSrc = new VisioAutomation.Core.SidSrc(-1,src);
            this.Value = value;
        }
    }
}