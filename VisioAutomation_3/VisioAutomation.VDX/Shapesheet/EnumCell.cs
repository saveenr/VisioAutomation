namespace VisioAutomation.VDX.ShapeSheet
{
    public class EnumCell<T> : CellScalar<T> where T : struct
    {
        private System.Func<T, int> EnumToInt;

        public EnumCell(System.Func<T, int> enum_to_int)
            : base(CellUnit.None)
        {
            this.EnumToInt = enum_to_int;
        }

        public override string GetResultString()
        {
            return string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0}", this.EnumToInt(this.Result.Value));
        }
    }
}