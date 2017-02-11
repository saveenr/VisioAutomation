using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal struct WriteRecord
    {
        private readonly SIDSRC _SIDSRC;
        private readonly SRC _SRC;

        internal readonly ValueLiteral Value;
        internal readonly CoordType Type;
        internal readonly IVisio.VisUnitCodes? UnitCode;

        public WriteRecord(SIDSRC sidsrc, ValueLiteral value, IVisio.VisUnitCodes? unitcode)
        {
            this._SIDSRC = sidsrc;
            this._SRC = new SRC();
            this.Value = value;
            this.Type = CoordType.SIDSRC;
            this.UnitCode = unitcode;
        }

        public WriteRecord(SRC src, ValueLiteral value, IVisio.VisUnitCodes? unitcode)
        {
            this._SIDSRC = new SIDSRC();
            this._SRC = src;
            this.Value = value;
            this.Type = CoordType.SRC;
            this.UnitCode = unitcode;
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