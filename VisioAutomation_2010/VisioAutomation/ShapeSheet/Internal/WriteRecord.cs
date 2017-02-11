using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal struct WriteRecord
    {
        private readonly SIDSRC _SIDSRC;

        internal readonly string CellValue;
        internal readonly CoordType CoordType;
        internal readonly IVisio.VisUnitCodes? UnitCode;

        public WriteRecord(SIDSRC sidsrc, string value, IVisio.VisUnitCodes? unitcode)
        {
            this._SIDSRC = sidsrc;
            this.CellValue = value;
            this.CoordType = CoordType.SIDSRC;
            this.UnitCode = unitcode;
        }

        public WriteRecord(SRC src, string value, IVisio.VisUnitCodes? unitcode)
        {
            this._SIDSRC = new SIDSRC(-1,src);
            this.CellValue = value;
            this.CoordType = CoordType.SRC;
            this.UnitCode = unitcode;
        }

        public SIDSRC SIDSRC
        {
            get
            {
                if (this.CoordType != CoordType.SIDSRC)
                {
                    throw new VisioAutomation.Exceptions.InternalAssertionException("Record does not contain a SIDSRC");
                }
                return _SIDSRC;
            }
        }

        public SRC SRC
        {
            get
            {
                if (this.CoordType != CoordType.SRC)
                {
                    throw new VisioAutomation.Exceptions.InternalAssertionException("Record does not contain a SRC");
                }
                return this._SIDSRC.SRC;
            }
        }
    }
}