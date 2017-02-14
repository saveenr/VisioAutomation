namespace VisioAutomation.ShapeSheet
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

        public override void AddSRC(SRC src)
        {
            this.Add(src);
        }
    }
}