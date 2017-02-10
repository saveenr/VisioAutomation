namespace VisioAutomation.ShapeSheet.Writers
{
    public struct WriteRecord<TValue>
    {
        private readonly SIDSRC SIDSRC;
        public readonly SRC SRC;
        public readonly TValue Value;
        public readonly CoordType Type;

        public WriteRecord(SIDSRC sidsrc, TValue value)
        {
            this.SIDSRC = sidsrc;
            this.SRC = new SRC();
            this.Value = value;
            this.Type = CoordType.SIDSRC;
        }

        public WriteRecord(SRC src, TValue value)
        {
            this.SIDSRC = new SIDSRC();
            this.SRC = src;
            this.Value = value;
            this.Type = CoordType.SRC;
        }

        public SIDSRC Sidsrc
        {
            get
            {
                if (this.Type != CoordType.SIDSRC)
                {
                    throw new System.ArgumentException();
                }
                return SIDSRC;
            }
        }

        public SRC Src
        {
            get
            {
                if (this.Type != CoordType.SRC)
                {
                    throw new System.ArgumentException();
                }
                return SRC;
            }
        }
    }
}