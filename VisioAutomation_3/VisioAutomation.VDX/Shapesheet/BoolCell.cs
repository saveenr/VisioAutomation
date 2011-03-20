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
            return (this.Result.Value) ? "1" : "0";
        }
    }
}