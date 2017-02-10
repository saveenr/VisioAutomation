using IVisio= Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public struct ResultValue
    {
        public readonly double ValueNumeric;
        public readonly string ValueString;
        public readonly IVisio.VisUnitCodes UnitCode;
        public readonly ResultType ResultType;

        internal ResultValue(double value,
            IVisio.VisUnitCodes unit_code)
        {
            this.UnitCode = unit_code;
            this.ValueNumeric = value;
            this.ValueString = null;
            this.ResultType = ResultType.ResultNumeric;
        }

        internal ResultValue(string value,
            IVisio.VisUnitCodes unit_code)
        {
            this.UnitCode = unit_code;
            this.ValueNumeric = 0.0;
            this.ValueString = value;
            this.ResultType = ResultType.ResultString;
        }

    }
}