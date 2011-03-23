namespace VisioAutomation.VDX.ShapeSheet
{
    public class BoolCell : CellScalar<bool>
    {
        public BoolCell()
            : base(CellUnit.None)
        {
        }

        public override string GetResultString()
        {
            if (!this.Result.HasValue)
            {
                throw new System.ArgumentException("BoolCell has no value");
            }

            return (this.Result.Value) ? "1" : "0";
        }
    }
}