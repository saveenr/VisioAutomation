namespace VisioAutomation.ShapeSheet.Writers
{
    public struct ResultValue
    {
        public readonly double ResultNumeric;
        public readonly string ResultString;
        public readonly Microsoft.Office.Interop.Visio.VisUnitCodes UnitCode;
        public readonly ResultType ResultType;

        internal ResultValue(double result,
            Microsoft.Office.Interop.Visio.VisUnitCodes unit_code)
        {
            this.UnitCode = unit_code;
            this.ResultNumeric = result;
            this.ResultString = null;
            this.ResultType = ResultType.ResultNumeric;
        }

        internal ResultValue(string result,
            Microsoft.Office.Interop.Visio.VisUnitCodes unit_code)
        {
            this.UnitCode = unit_code;
            this.ResultNumeric = 0.0;
            this.ResultString = result;
            this.ResultType = ResultType.ResultString;
        }

    }
}