namespace VisioAutomation.ShapeSheet.Writers
{
    public struct WriteRecord<TValue>
    {
        private readonly SIDSRC _SIDSRC;
        private readonly SRC _SRC;

        public readonly TValue Value;
        public readonly CoordType Type;

        public WriteRecord(SIDSRC sidsrc, TValue value)
        {
            this._SIDSRC = sidsrc;
            this._SRC = new SRC();
            this.Value = value;
            this.Type = CoordType.SIDSRC;
        }

        public WriteRecord(SRC src, TValue value)
        {
            this._SIDSRC = new SIDSRC();
            this._SRC = src;
            this.Value = value;
            this.Type = CoordType.SRC;
        }

        public SIDSRC SIDSRC
        {
            get
            {
                if (this.Type != CoordType.SIDSRC)
                {
                    throw new System.ArgumentException();
                }
                return _SIDSRC;
            }
        }

        public SRC SRC
        {
            get
            {
                if (this.Type != CoordType.SRC)
                {
                    throw new System.ArgumentException();
                }
                return _SRC;
            }
        }
    }
}