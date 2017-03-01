namespace VisioAutomation.ShapeSheet
{
    class WriterCollection_SRC
    {
        public VisioAutomation.ShapeSheet.Streams.SRCStreamBuilder StreamBuilder;
        public VisioAutomation.ShapeSheet.Internal.ObjectArrayBuilder<string> ValuesBuilder;


        public WriterCollection_SRC()
        {

            this.StreamBuilder = new VisioAutomation.ShapeSheet.Streams.SRCStreamBuilder();
            this.ValuesBuilder = new VisioAutomation.ShapeSheet.Internal.ObjectArrayBuilder<string>();
        }

        public void Clear()
        {
            this.StreamBuilder.Clear();
            this.ValuesBuilder.Clear();
        }


        public void Add(SRC src, string value)
        {
            this.StreamBuilder.Add(src);
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

        public object[] BuildUnitCodes()
        {
            return null;
        }

        public int Count => this.StreamBuilder.Count();
    }
}