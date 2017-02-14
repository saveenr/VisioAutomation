namespace VisioAutomation.ShapeSheet
{
    public class SRCStreamBuilder : StreamBuilder<SRC>
    {
        public SRCStreamBuilder() : base()
        {
            
        }

        public SRCStreamBuilder(int capacity) : base(capacity)
        {

        }

        protected override ShapeSheetStream build_stream()
        {
            return new SIDSRCStream(SRC.ToStream(this.items));
        }

        internal override void AddSRC(SRC src)
        {
            this.Add(src);
        }
    }

    public class ShapeSheetStream
    {
        internal short[] short_array;

        internal ShapeSheetStream(short[] a)
        {
            this.short_array = a;
        }

        public bool IsEmpty => this.short_array.Length == 0;
    }

    public class SRCStream : ShapeSheetStream
    {

        internal SRCStream(short[] a) : base(a)
        {
        }
    }

    public class SIDSRCStream : ShapeSheetStream
    {

        internal SIDSRCStream(short[] a) : base(a)
        {
        }
    }

}