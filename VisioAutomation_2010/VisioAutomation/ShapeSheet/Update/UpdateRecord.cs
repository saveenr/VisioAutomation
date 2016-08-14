
namespace VisioAutomation.ShapeSheet.Update
{
    public struct UpdateRecord
    {
        public readonly SIDSRC SIDSRC;
        public readonly string Formula;
        public readonly double ResultNumeric;
        public readonly string ResultString;
        public readonly Microsoft.Office.Interop.Visio.VisUnitCodes UnitCode;
        public readonly UpdateType UpdateType;
        public readonly StreamType StreamType;

        internal UpdateRecord(StreamType streamtype, SIDSRC sidsrc, string formula)
        {
            this.SIDSRC = sidsrc;
            this.Formula = formula;
            this.ResultNumeric = 0.0;
            this.ResultString = null;
            this.UnitCode = Microsoft.Office.Interop.Visio.VisUnitCodes.visNumber;
            this.UpdateType = UpdateType.Formula;
            this.StreamType = streamtype;
        }

        internal UpdateRecord(StreamType streamtype, SIDSRC sidsrc, double result, Microsoft.Office.Interop.Visio.VisUnitCodes unit_code)
        {
            this.SIDSRC = sidsrc;
            this.Formula = null;
            this.UnitCode = unit_code;
            this.ResultNumeric = result;
            this.ResultString = null;
            this.UpdateType = UpdateType.ResultNumeric;
            this.StreamType = streamtype;
        }

        internal UpdateRecord(StreamType streamtype, SIDSRC sidsrc, string result, Microsoft.Office.Interop.Visio.VisUnitCodes unit_code)
        {
            this.SIDSRC = sidsrc;
            this.Formula = null;
            this.UnitCode = unit_code;
            this.ResultNumeric = 0.0;
            this.ResultString = result;
            this.UpdateType = UpdateType.ResultString;
            this.StreamType = streamtype;
        }

    }
}