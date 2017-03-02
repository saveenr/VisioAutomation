namespace VisioAutomation.ShapeSheet.Internal
{
    class WriterCollection_Src
    {
        public VisioAutomation.ShapeSheet.Streams.SrcStreamBuilder StreamBuilder;
        public VisioAutomation.ShapeSheet.Internal.ObjectArrayBuilder<string> ValuesBuilder;


        public WriterCollection_Src()
        {

            this.StreamBuilder = new VisioAutomation.ShapeSheet.Streams.SrcStreamBuilder();
            this.ValuesBuilder = new VisioAutomation.ShapeSheet.Internal.ObjectArrayBuilder<string>();
        }

        public void Clear()
        {
            this.StreamBuilder.Clear();
            this.ValuesBuilder.Clear();
        }


        public void Add(Src src, string value)
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

        public int Count => this.StreamBuilder.Count();
    }
}