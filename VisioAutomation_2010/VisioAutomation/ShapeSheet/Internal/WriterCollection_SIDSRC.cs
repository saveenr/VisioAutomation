namespace VisioAutomation.ShapeSheet.Internal
{
    class WriterCollection_SidSrc
    {
        private VisioAutomation.ShapeSheet.Streams.SidSrcStreamBuilder StreamBuilder;
        private VisioAutomation.ShapeSheet.Internal.ObjectArrayBuilder<string> ValuesBuilder;

        public WriterCollection_SidSrc()
        {
            this.StreamBuilder = new VisioAutomation.ShapeSheet.Streams.SidSrcStreamBuilder();
            this.ValuesBuilder = new VisioAutomation.ShapeSheet.Internal.ObjectArrayBuilder<string>();
        }

        public void Clear()
        {
            this.StreamBuilder.Clear();
            this.ValuesBuilder.Clear();
        }

        public void Add(SidSrc sidsrc, string value)
        {
            this.StreamBuilder.Add(sidsrc);
            this.ValuesBuilder.Add(value);
        }

        public void Add(SidSrc sidsrc, string value, Microsoft.Office.Interop.Visio.VisUnitCodes unit_code)
        {
            this.StreamBuilder.Add(sidsrc);
            this.ValuesBuilder.Add(value);
        }


        public short[] BuildStream()
        {
            return this.StreamBuilder.ToStream();
        }

        public object[] BuildValues()
        {
            return this.ValuesBuilder.ToObjectArray();
        }

        public int Count => this.StreamBuilder.Count;

    }
}