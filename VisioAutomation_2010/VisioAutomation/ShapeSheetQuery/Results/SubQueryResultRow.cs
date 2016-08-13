namespace VisioAutomation.ShapeSheetQuery.Results
{
    public struct SubQueryResultRow<T> 
    {
        public readonly T[] Cells;

        internal SubQueryResultRow(T[] c)
        {
            this.Cells = c;
        }
    }
}