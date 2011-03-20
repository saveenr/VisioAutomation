namespace VisioAutomation.ShapeSheet.Update
{
    public struct FormulaItem<TStream> where TStream : struct
    {
        public readonly TStream StreamItem;
        public readonly string Formula;

        public FormulaItem(TStream streamitem, string formula)
        {
            this.StreamItem = streamitem;
            this.Formula = formula;
        }
    }
}