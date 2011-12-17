namespace VisioAutomation.VDX.ShapeSheet
{
    public class DoubleCell : CellScalar<double>
    {
        public DoubleCell()
            : base(CellUnit.None)
        {
        }

        public DoubleCell(double value)
            : base(CellUnit.None)
        {
            this.Result = value;
        }
    }
}