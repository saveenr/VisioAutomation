namespace VisioAutomation.VDX.ShapeSheet
{
    public class ColorCell : CellScalar<int>
    {
        public ColorCell()
            : base(CellUnit.None)
        {
        }

        public override string GetResultString()
        {
            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            return string.Format(invariant_culture, "#{0:X}", this.Result);
        }
    }
}