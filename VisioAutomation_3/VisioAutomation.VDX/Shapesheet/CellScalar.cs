namespace VisioAutomation.VDX.ShapeSheet
{
    public class CellScalar<T> : CellBase where T : struct
    {
        public T? Result;

        public CellScalar() :
            base()
        {
        }

        public CellScalar(CellUnit unit)
            : base(unit)
        {
        }

        public override string GetResultString()
        {
            return string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0}", this.Result);
        }

        public override bool HasResult
        {
            get { return this.Result.HasValue; }
        }
    }
}