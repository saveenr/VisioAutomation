namespace VisioAutomation.VDX.ShapeSheet
{
    public class PointCell : CellScalar<double>
    {
        public PointCell()
            : base(CellUnit.Point)
        {
        }

        public override string GetResultString()
        {
            var invariant_culture = System.Globalization.CultureInfo.InvariantCulture;
            var inches = Converter.PointsToInches(this.Result.Value);
            return string.Format(invariant_culture, "{0}", inches);
        }
    }
}