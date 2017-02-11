using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal struct WriteRecord
    {
        private readonly SIDSRC _SIDSRC;
        private readonly SRC _SRC;

        internal readonly string CellValue;
        internal readonly CoordType Type;
        internal readonly IVisio.VisUnitCodes? UnitCode;

        public WriteRecord(SIDSRC sidsrc, string value, IVisio.VisUnitCodes? unitcode)
        {
            this._SIDSRC = sidsrc;
            this._SRC = new SRC();
            this.CellValue = value;
            this.Type = CoordType.SIDSRC;
            this.UnitCode = unitcode;
        }

        public WriteRecord(SRC src, string value, IVisio.VisUnitCodes? unitcode)
        {
            this._SIDSRC = new SIDSRC();
            this._SRC = src;
            this.CellValue = value;
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