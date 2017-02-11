using IVisio= Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet.Writers
{
    public struct ResultValue
    {
        public readonly string ValueString;
        public readonly IVisio.VisUnitCodes UnitCode;

        internal ResultValue(double value,
            IVisio.VisUnitCodes unit_code)
        {
            this.UnitCode = unit_code;
            this.ValueString = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        internal ResultValue(string value,
            IVisio.VisUnitCodes unit_code)
        {
            this.UnitCode = unit_code;
            this.ValueString = value;
        }
    }
}