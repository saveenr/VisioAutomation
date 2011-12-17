using System;

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
            if (!this.Result.HasValue)
            {
                throw new ArgumentException("EnumCell does not have a value");
            }

            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            var the_int = this.EnumToInt(this.Result.Value);
            return string.Format(invariant_culture, "{0}", the_int);
        }
    }
}