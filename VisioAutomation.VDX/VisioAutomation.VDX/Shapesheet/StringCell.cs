namespace VisioAutomation.VDX.ShapeSheet
{
    public class StringCell : CellBase
    {
        public string Result;

        public override string GetResultString()
        {
            return this.Result;
        }

        public override bool HasResult
        {
            get { return this.Result != null; }
        }
    }
}