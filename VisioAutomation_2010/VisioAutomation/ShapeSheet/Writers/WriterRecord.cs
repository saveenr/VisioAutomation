
namespace VisioAutomation.ShapeSheet.Writers
{
    public struct WriterRecord<T>
    {
        public readonly T StreamItem;
        public readonly string Formula;
        public readonly double ResultNumeric;
        public readonly string ResultString;
        public readonly Microsoft.Office.Interop.Visio.VisUnitCodes UnitCode;
        public readonly UpdateType UpdateType;

        internal WriterRecord(T stream_item, string formula)
        {
            this.StreamItem = stream_item;
            this.Formula = formula;
            this.ResultNumeric = 0.0;
            this.ResultString = null;
            this.UnitCode = Microsoft.Office.Interop.Visio.VisUnitCodes.visNumber;
            this.UpdateType = UpdateType.Formula;
        }

        internal WriterRecord(T stream_item, double result, Microsoft.Office.Interop.Visio.VisUnitCodes unit_code)
        {
            this.StreamItem = stream_item;
            this.Formula = null;
            this.UnitCode = unit_code;
            this.ResultNumeric = result;
            this.ResultString = null;
            this.UpdateType = UpdateType.ResultNumeric;
        }

        internal WriterRecord(T stream_item, string result, Microsoft.Office.Interop.Visio.VisUnitCodes unit_code)
        {
            this.StreamItem = stream_item;
            this.Formula = null;
            this.UnitCode = unit_code;
            this.ResultNumeric = 0.0;
            this.ResultString = result;
            this.UpdateType = UpdateType.ResultString;
        }

    }
}