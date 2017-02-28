namespace VisioAutomation.ShapeSheet
{
    class WriterCollection_SRC
    {
        public SRCStreamBuilder StreamBuilder;
        public ShapeSheetObjectArrayBuilder<string> ValuesBuilder;


        public WriterCollection_SRC(bool include_unit_codes)
        {

            this.StreamBuilder = new SRCStreamBuilder();
            this.ValuesBuilder = new ShapeSheetObjectArrayBuilder<string>();
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