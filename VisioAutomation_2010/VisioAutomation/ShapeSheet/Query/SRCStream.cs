namespace VisioAutomation.ShapeSheet.Query
{
    public class SRCStream : StreamBase<SRC>
    {
        public SRCStream() : base()
        {
            
        }

        public SRCStream(int capacity) : base(capacity)
        {

        }

        protected override short[] get_stream()
        {
            return SRC.ToStream(this.items);
        }
    }
}